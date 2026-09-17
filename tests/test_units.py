import pandas as pd
import pytest

from voc import pii, sentiment
from voc.clean import Cleaner, junk_reason
from voc.codebook import Codebook
from voc.questions import answer_kind, detect_questions, role_for
from voc.textfix import decode_bytes, fix_cell

CB = Codebook.load()


# --- text repair -----------------------------------------------------------------
def test_decode_mac_roman_export():
    raw = "¿Cómo te enteraste de nosotros?,Total\nSí,1\n".encode("mac_roman")
    text, enc = decode_bytes(raw)
    assert enc == "mac_roman"
    assert text.startswith("¿Cómo te enteraste")


def test_decode_utf8_passthrough():
    text, enc = decode_bytes("¿Qué?".encode("utf-8"))
    assert (text, enc) == ("¿Qué?", "utf-8")


@pytest.mark.parametrize("broken,fixed", [
    ("Cu√©ntanos, ¬øcu√°l fue la raz√≥n", "Cuéntanos, ¿cuál fue la razón"),
    ("¬°Hola!   ¿Tienes dudas?", "¡Hola! ¿Tienes dudas?"),
])
def test_fix_cell_repairs_double_decoding(broken, fixed):
    assert fix_cell(broken) == fixed


# --- privacy -------------------------------------------------------------------
@pytest.mark.parametrize("header", ["Full Name", "Email", "City", "Province", "ipAddress", "userAgent",
                                    "reviewer_name", "Survey Metadata: url", "original-referrer", "First Name"])
def test_pii_columns_detected(header):
    assert pii.is_pii_column(header)


@pytest.mark.parametrize("header", ["Line Items", "rating", "Started At", "¿Qué te hizo comprar?"])
def test_non_pii_columns_kept(header):
    assert not pii.is_pii_column(header)


def test_scrub_text_masks_contact_details():
    text = "Escríbanme a ana.demo@example.com o al 81 1234 5678, pedido #123456, sigo a @marca_mx https://x.co/a"
    out = pii.scrub_text(text)
    for secret in ("ana.demo@example.com", "1234 5678", "123456", "@marca_mx", "https://x.co/a"):
        assert secret not in out
    assert "[email]" in out and "[phone]" in out and "[order]" in out and "[handle]" in out and "[url]" in out


def test_spend_band():
    assert pii.spend_band(739) == "500-999"
    assert pii.spend_band("17598.5") == "5000+"
    assert pii.spend_band(None) == "unknown"


# --- question roles --------------------------------------------------------------
@pytest.mark.parametrize("header,role", [
    ("¿Que te causaba incertidumbre antes de comprar con Nosotros?", "hesitation"),
    ("¿Cual era tu miedo más grande al comprar con nosotros ?", "hesitation"),
    ("¿Qué te hizo comprar X antes que otra marca?", "choice_driver"),
    ("Cuéntanos, ¿cuál fue la razón principal por la que decidiste comprar?", "purchase_reason"),
    ("¿Qué problemas estas buscando resolver al comprar X?", "purchase_reason"),
    ("¿Cómo nos conociste por primera vez?", "discovery"),
    ("¿Y qué te trajo a nuestra página hoy?", "visit_trigger"),
    ("¿En dónde te enteraste del anuncio de nosotros?", "ad_location"),
    ("¿Tenías en mente alguna otra marca justo antes de decidirte?", "competitor"),
    ("Slide: ¡Hola! ¿Tienes alguna inquietud o duda que te dificulte comprar? | Id: 64d", "pdp_gate"),
    ("Cuéntanos, ¿cuál es?", "pdp_blocker"),
    ("body", "review"),
    ("Q3", "open_feedback"),
])
def test_role_for(header, role):
    assert role_for(header) == role


def test_answer_kind_gate_and_choice():
    assert answer_kind(pd.Series(["Sí", "No"] * 20), "pdp_gate") == "gate"
    assert answer_kind(pd.Series(["Instagram", "Facebook", "TikTok"] * 20), "discovery") == "choice"
    assert answer_kind(pd.Series([f"respuesta libre {i}" for i in range(40)]), "hesitation") == "text"


def test_single_open_gate_question_becomes_blocker():
    df = pd.DataFrame({"¡Hola! ¿Tienes alguna inquietud o duda?": [f"duda {i}" for i in range(30)]})
    (q,) = detect_questions(df)
    assert (q.role, q.kind) == ("pdp_blocker", "text")


def test_pasted_duplicate_question_is_dropped():
    answers = [f"respuesta {i}" for i in range(30)]
    df = pd.DataFrame({"¿Qué te preocupaba?": answers, "Slide: ¿Qué te preocupaba?": answers})
    assert len(detect_questions(df)) == 1


# --- codebook --------------------------------------------------------------------
@pytest.mark.parametrize("answer,theme", [
    ("Que no llegara el pedido", "shipping_delivery"),
    ("que no funcione", "efficacy_results"),
    ("Que fuera una estafa", "trust_legitimacy"),
    ("Ninguno 😊", "no_concern"),
    ("nada en especial", "no_concern"),
    ("¿Tienen tabla de medidas?", "size_fit"),
    ("¿Aceptan pago en OXXO?", "payment_checkout"),
    ("Me lo recomendó una amiga", "social_proof"),
    ("Tengo SOP y quiero regular mi ciclo", "health_goal"),
    ("Cuando vuelven a tener el modelo azul", "stock_availability"),
    ("Saber cómo me ayudan", "product_info"),
    ("costura chueca", "quality"),
    ("zzz sin relación", "other"),
])
def test_codebook_primary_theme(answer, theme):
    assert CB.match(answer)[0] == theme


def test_codebook_orders_themes_by_position():
    assert CB.match("Excelente precio pero tardó en llegar") == ["price_value", "shipping_delivery"]


def test_side_effects_are_not_efficacy():
    assert CB.match("efectos secundarios")[0] == "ingredients_safety"


def test_extra_keywords(tmp_path):
    extra = tmp_path / "extra.json"
    extra.write_text('{"social_proof": ["Nutrióloga Demo"]}', encoding="utf-8")
    cb = Codebook.load(extra_keywords=extra)
    assert cb.match("sigo a la nutrióloga demo")[0] == "social_proof"


def test_extra_keywords_unknown_theme(tmp_path):
    extra = tmp_path / "extra.json"
    extra.write_text('{"not_a_theme": ["x"]}', encoding="utf-8")
    with pytest.raises(ValueError):
        Codebook.load(extra_keywords=extra)


# --- sentiment -------------------------------------------------------------------
@pytest.mark.parametrize("text,label", [
    ("Me encantó, excelente producto", "Positive"),
    ("No me gustó para nada", "Negative"),
    ("Sin ningún problema, llegó bien", "Positive"),
    ("Huele rico pero dura poco", "Mixed"),
    ("Mala calidad", "Negative"),
    ("Es el de 100 ml", "Neutral"),
    ("Lo amo 😍", "Positive"),
    ("Great quality, but shipping was late", "Mixed"),
])
def test_sentiment(text, label):
    assert sentiment.classify(text) == label


def test_rating_sentiment():
    assert sentiment.from_rating(5) == "Positive"
    assert sentiment.from_rating("2") == "Negative"
    assert sentiment.from_rating(None) is None


# --- cleaning --------------------------------------------------------------------
@pytest.mark.parametrize("answer,reason", [
    ("", "empty"), (".", "no_letters"), ("5", "no_letters"), ("x", "single_letter"),
    ("test", "test_or_greeting"), ("Hola", "test_or_greeting"), ("jjjjjj", "repeated_character"),
    ("asdfgh", "keyboard_mash"), ("sdkjfhskdjfh", "keyboard_mash"),
])
def test_junk_detected(answer, reason):
    assert junk_reason(answer) == reason


@pytest.mark.parametrize("answer", ["No", "nada", "N/A", "10/10", "😍😍", "Fabulososssss!!!", "Probióticos", "property"])
def test_not_junk(answer):
    assert junk_reason(answer) is None


def test_cleaner_fixes_typos_and_shorthand():
    cleaner = Cleaner().fit(["productos"] * 10)
    assert cleaner.clean("q buena calidsd").text == "que buena calidad"
    assert cleaner.clean("Ninguo").text == "Ninguno"
    assert cleaner.clean("los prodcutos xq funcionan").text == "los productos porque funcionan"
    assert cleaner.clean("me encantaaaa").text == "me encanta"


def test_cleaner_leaves_real_words_alone():
    cleaner = Cleaner().fit(["pero"] * 20 + ["precio"] * 50)
    for sentence in ("padecia de acné", "tomandolos diario", "batallaba mucho", "hago ejercicio", "buen producto"):
        assert cleaner.clean(sentence).text == sentence


def test_cleaner_learns_repeated_misspelling():
    cleaner = Cleaner().fit(["envio"] * 40 + ["evnio"] * 3)
    assert cleaner.clean("evnio rapido").text == "envio rapido"


# --- multi-select detection -------------------------------------------------------
def test_multi_select_column_is_not_free_text():
    options = ["Precio Justo", "Entrega Rapida", "Servicio al Cliente", "Garantia de por Vida"]
    answers = [", ".join(options[: 1 + i % 3]) for i in range(40)]
    assert answer_kind(pd.Series(answers), "purchase_reason") == "multi"


def test_free_text_is_not_mistaken_for_multi_select():
    answers = [f"me preocupaba que no llegara el pedido {i}, la verdad" for i in range(40)]
    assert answer_kind(pd.Series(answers), "hesitation") == "text"


@pytest.mark.parametrize("answer,theme", [
    ("Como puedo comprar con mis puntos acumulados", "loyalty_rewards"),
    ("Facilidad en la compra y en la explicación", "site_usability"),
    ("La talla 4 aparece con una diagonal, no entiendo", "site_usability"),
    ("Quiero probar la marca", "first_time_trial"),
    ("Nunca había comprado aquí", "first_time_trial"),
    ("Ya he comprado antes y todo bien", "repeat_loyalty"),
    ("La seguridad de comprar en línea", "trust_legitimacy"),
])
def test_themes_added_after_diagnosing_real_exports(answer, theme):
    assert theme in CB.match(answer)

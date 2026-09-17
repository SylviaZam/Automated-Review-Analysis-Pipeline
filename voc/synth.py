"""Synthetic "twin" datasets for the public demo.

The real client exports cannot be published. These files copy their *shape*
(question wording, column layout, bilingual answers, typos, missing accents,
fake PII columns that ingest must drop) but every answer is written from
templates and every person is fictional. Theme proportions are invented and
are not client results.

Because the generator knows which theme each answer was written for, it also
emits a gold file for the evaluation harness.
"""
from __future__ import annotations

import random
import unicodedata
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Templates: {theme_id: [answers]} per question. Written as customers write,
# not from the codebook keyword list, so some answers are deliberately hard.
# ---------------------------------------------------------------------------
WELLNESS_PROBLEM = {
    "health_goal": [
        "Tengo SOP y quiero regular mi ciclo", "Dolor de rodillas", "Resistencia a la insulina",
        "Bajar de peso despues del embarazo", "Mejorar mi digestión, vivo inflamada", "Dormir mejor",
        "Me duelen mucho las articulaciones", "Acné hormonal", "Tener más energía en el día",
        "Controlar la ansiedad", "Se me cae mucho el cabello", "Colitis", "Mejorar mi salud en general",
        "joint pain", "I want to sleep better", "Cansancio todo el tiempo",
    ],
    "efficacy_results": ["Quiero ver resultados reales esta vez", "Probar algo que sí funcione, ya intenté de todo"],
    "no_concern": ["Ninguno", "nada en especial"],
    "other": ["Es un regalo para mi mamá", "Curiosidad", "La verdad no sé"],
}
WELLNESS_PROBLEM_W = {"health_goal": 78, "efficacy_results": 8, "no_concern": 6, "other": 8}

WELLNESS_DRIVER = {
    "ingredients_safety": [
        "Que los ingredientes son naturales", "No tiene azúcar ni cosas raras", "Está respaldado por ciencia",
        "Me dio confianza que fuera orgánico", "Clean ingredients", "Sin químicos",
    ],
    "social_proof": [
        "Me lo recomendó una amiga", "Por las reseñas", "Sigo a la nutrióloga en Instagram y la recomienda",
        "Una compañera de trabajo los toma", "Vi muchos testimonios", "Mi hermana los usa y le fue muy bien",
        "Por la comunidad que tienen", "my friend recommended it",
    ],
    "efficacy_results": [
        "Porque sí funcionan", "A mi prima le dio resultado", "Ya había probado otros y no me sirvieron",
        "Se nota la diferencia", "Me ayudó muchísimo con la inflamación",
    ],
    "price_value": ["El precio", "Estaba en promoción", "Más barato que en la farmacia", "Tenían 20% de descuento"],
    "repeat_loyalty": ["Ya lo había comprado antes", "Siempre les compro", "Es mi segunda vez"],
    "brand_values": ["Es una marca mexicana", "Me gusta apoyar negocios locales"],
    "design_style": ["El empaque está muy bonito"],
    "other": ["No sé, me llamó la atención", "Porque sí", "Me urgía"],
}
WELLNESS_DRIVER_W = {"ingredients_safety": 22, "social_proof": 26, "efficacy_results": 18, "price_value": 9,
                     "repeat_loyalty": 9, "brand_values": 4, "design_style": 2, "other": 10}

WELLNESS_HESITATION = {
    "efficacy_results": [
        "Que no funcione", "Que no me sirva como a los demás", "Gastar y no ver ningún resultado",
        "Que fuera puro marketing", "Que no hiciera efecto", "That it wouldn't work", "que no funcionaran",
    ],
    "shipping_delivery": [
        "Que no llegara el pedido", "Que tardara mucho el envío", "Que se perdiera el paquete",
        "Que no llegue a mi ciudad", "El tiempo de entrega", "que no me llegara",
    ],
    "trust_legitimacy": [
        "Que fuera fraude", "No conocía la marca", "Comprar en una página que no conozco",
        "Que fuera una estafa", "Dar mis datos de tarjeta en una tienda nueva", "Is this brand legit?",
    ],
    "no_concern": ["Ninguno", "Nada", "ninguna", "No tenía ningún miedo", "Ninguno 😊", "none"],
    "price_value": ["El precio, está un poco caro", "Que saliera muy caro con el envío"],
    "ingredients_safety": ["Que me cayera mal", "Efectos secundarios", "Soy alérgica a varias cosas"],
    "other": ["Que no me gustara el sabor", "No sé", "Que mi esposo se enojara jaja"],
}
WELLNESS_HESITATION_W = {"efficacy_results": 32, "shipping_delivery": 20, "trust_legitimacy": 13, "no_concern": 19,
                         "price_value": 6, "ingredients_safety": 5, "other": 5}

DISCOVERY = {"Instagram": 38, "Me lo recomendaron": 24, "TikTok": 12, "Facebook": 10, "Google": 9, "Otro": 7}
VISIT = {"Por mi propia cuenta": 41, "Me acordé de ustedes": 17, "Vi un anuncio": 16,
         "Alguien me habló de ti": 13, "Recibí un e-mail": 8, "Estaba buscando algo en Google": 5}
AD = {"Instagram": 64, "Facebook": 21, "TikTok": 11, "YouTube": 4}

PDP_BLOCKER = {
    "size_fit": [
        "Talla", "¿Cómo sé qué talla pedir?", "Si la talla viene correcta", "Las tallas vienen chicas?",
        "Tienen tabla de medidas?", "que talla le queda a un niño de 4 años", "does it run small?",
    ],
    "shipping_delivery": [
        "¿Hacen envíos a USA?", "Tiempo de entrega", "Cuánto tarda en llegar a Mérida", "¿El envío es gratis?",
        "Envían a todo México?", "cuanto cuesta el envio",
    ],
    "stock_availability": [
        "¿Cuándo tendrán otra vez el modelo de estrellas?", "Ya no hay en talla 6?", "Está agotado el que quiero",
        "¿Van a resurtir?", "Cuando vuelven a tener el pijama azul",
    ],
    "payment_checkout": [
        "¿Aceptan pago en OXXO?", "Cómo realizo el pedido", "Puedo pagar con PayPal?", "¿Tienen meses sin intereses?",
        "No me deja pagar con mi tarjeta",
    ],
    "customer_service_returns": [
        "Qué tan complicado es hacer un cambio", "Si no le queda, ¿lo puedo cambiar?", "¿Tienen devoluciones?",
        "Nadie me contesta el WhatsApp",
    ],
    "product_info": ["¿De qué material es?", "Se puede meter a la secadora?", "¿Qué incluye el set?"],
    "price_value": ["¿Tienen algún descuento?", "Está caro"],
    "other": ["Hola", "Ninguna gracias", "Son lindos", "ok"],
}
PDP_BLOCKER_W = {"size_fit": 30, "shipping_delivery": 22, "stock_availability": 14, "payment_checkout": 10,
                 "customer_service_returns": 9, "product_info": 7, "price_value": 3, "other": 5}

REVIEW_POS = {
    "quality": [
        "Excelente calidad, la fragancia dura todo el día", "Huele delicioso y dura muchísimo",
        "Muy buena fijación, igualito al original", "Me encantó, el aroma se mantiene horas", "Great quality, lasts all day",
    ],
    "price_value": ["Excelente precio para lo que es", "Muy buen perfume y súper barato", "Vale la pena, buen precio"],
    "shipping_delivery": ["Llegó rapidísimo y bien empacado", "Todo perfecto, el envío muy rápido", "Llegó antes de lo esperado, gracias"],
    "design_style": ["La presentación está hermosa", "Me gustó mucho el frasco, muy bonito"],
    "social_proof": ["Lo compré por las reseñas y no me arrepiento", "Me lo recomendó mi amiga y me encantó"],
    "other": ["Me encanta 😍", "10/10", "Mi favorito", "Perfecto", "Lo amo", "Súper recomendado"],
}
REVIEW_POS_W = {"quality": 38, "price_value": 14, "shipping_delivery": 14, "design_style": 6, "social_proof": 6, "other": 22}

REVIEW_NEG = {
    "quality": ["No dura nada, a la hora ya no huele", "El olor no se parece al original", "Mala fijación, no lo recomiendo"],
    "shipping_delivery": ["Tardó tres semanas en llegar", "Nunca llegó mi pedido", "Llegó el frasco roto"],
    "customer_service_returns": ["Pedí un cambio y nadie me respondió", "Pésima atención al cliente"],
    "other": ["No me gustó", "Decepcionada", "Meh"],
}
REVIEW_NEG_W = {"quality": 45, "shipping_delivery": 30, "customer_service_returns": 12, "other": 13}

REVIEW_MIXED = {
    "quality": ["Huele rico pero dura poco", "Buen aroma aunque la fijación es regular"],
    "shipping_delivery": ["Me encantó el perfume pero tardó mucho en llegar", "Excelente producto, solo que el envío fue lento"],
}
REVIEW_NEUTRAL = {
    "quality": ["Es parecido al original", "Aroma dulce, fijación normal"],
    "other": ["Es para regalo, todavía no lo abro", "Es el de 100 ml"],
}

FIRST = ["Ana", "Luis", "María", "Sofía", "Carlos", "Valeria", "Diego", "Fernanda", "Jorge", "Daniela", "Lucía", "Pablo"]
LAST = ["Demo", "Ejemplo", "Prueba", "Ficticio", "Muestra"]
CITIES = ["Ciudad Ejemplo", "Pueblo Demo", "Villa Prueba"]
WELLNESS_PRODUCTS = ["Colágeno Rosa", "Inositol Balance", "Magnesio Noche", "Omega Calma", "Proteína Vainilla", "Probiótico Diario"]
APPAREL_PRODUCTS = ["pijama-estrellas", "pijama-azul", "set-dinosaurio", "vestido-nube", "sudadera-luna"]
FRAGRANCE_PRODUCTS = ["eau-de-parfum-ambar-100ml", "eau-de-parfum-flor-50ml", "body-mist-coco", "eau-de-toilette-cedro-100ml"]


def _strip_accents(text: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch))


def _noise(rng: random.Random, text: str) -> str:
    if rng.random() < 0.35:
        text = _strip_accents(text)
    if rng.random() < 0.25:
        text = text.lower()
    if rng.random() < 0.15:
        text = text.rstrip(".") + "."
    if rng.random() < 0.08 and len(text) > 6:
        i = rng.randrange(1, len(text) - 1)
        text = text[:i] + text[i + 1:]  # dropped letter
    return text


def _pick(rng: random.Random, templates: dict, weights: dict) -> tuple[str, str]:
    theme = rng.choices(list(weights), weights=list(weights.values()))[0]
    return theme, rng.choice(templates[theme])


def _weighted(rng: random.Random, options: dict) -> str:
    return rng.choices(list(options), weights=list(options.values()))[0]


def _person(rng: random.Random, i: int) -> dict:
    first, last = rng.choice(FIRST), rng.choice(LAST)
    return {"Full Name": f"{first} {last}", "Email": f"demo+{i:04d}@example.com", "City": rng.choice(CITIES)}


def wellness_post_purchase(rng: random.Random, n: int = 1200):
    rows, gold = [], []
    q_problem = "¿Cuál es el problema más importante que esperas resolver con nuestros productos?"
    q_driver = "¿Qué te hizo comprar con nosotros antes que otra marca?"
    q_hes = "¿Qué te causaba incertidumbre antes de comprar con nosotros?"
    for i in range(1, n + 1):
        row = _person(rng, i)
        row["Total Spent"] = round(rng.uniform(350, 4200), 2)
        row["Line Items"] = ", ".join(rng.sample(WELLNESS_PRODUCTS, rng.choice([1, 1, 2, 3])))
        row["Started At"] = f"2024-{rng.randint(1, 6):02d}-{rng.randint(1, 28):02d} {rng.randint(8, 23):02d}:{rng.randint(0, 59):02d}"
        for q, templates, weights, role, skip in (
            (q_problem, WELLNESS_PROBLEM, WELLNESS_PROBLEM_W, "purchase_reason", 0.10),
            (q_driver, WELLNESS_DRIVER, WELLNESS_DRIVER_W, "choice_driver", 0.15),
            (q_hes, WELLNESS_HESITATION, WELLNESS_HESITATION_W, "hesitation", 0.12),
        ):
            if rng.random() < skip:
                row[q] = ""
                continue
            theme, text = _pick(rng, templates, weights)
            text = _noise(rng, text)
            row[q] = text
            gold.append({"role": role, "question": q, "answer": text, "gold_theme": theme, "gold_sentiment": ""})
        row["¿Cómo nos conociste por primera vez?"] = _weighted(rng, DISCOVERY)
        row["¿Y qué te trajo a nuestra página hoy?"] = _weighted(rng, VISIT)
        row["¿Dónde estaba el anuncio?"] = _weighted(rng, AD) if row["¿Y qué te trajo a nuestra página hoy?"] == "Vi un anuncio" else ""
        rows.append(row)
    return pd.DataFrame(rows), gold


def apparel_pdp(rng: random.Random, n: int = 900):
    rows, gold = [], []
    q_gate = "¡Hola! ¿Tienes alguna inquietud o duda que te dificulte comprar?"
    q_open = "Cuéntanos, ¿cuál es?"
    for i in range(1, n + 1):
        yes = rng.random() < 0.18
        text = ""
        if yes and rng.random() < 0.9:
            theme, text = _pick(rng, PDP_BLOCKER, PDP_BLOCKER_W)
            text = _noise(rng, text)
            gold.append({"role": "pdp_blocker", "question": q_open, "answer": text, "gold_theme": theme, "gold_sentiment": ""})
        rows.append({
            "Created Date": f"08/{rng.randint(1, 31):02d}/2024 {rng.randint(1, 12):02d}:{rng.randint(0, 59):02d}pm",
            "handle": rng.choice(APPAREL_PRODUCTS),
            q_gate: "Sí" if yes else "No",
            q_open: text,
            "ipAddress": f"203.0.113.{rng.randint(1, 254)}",
        })
    return pd.DataFrame(rows), gold


def fragrance_reviews(rng: random.Random, n: int = 700):
    rows, gold = [], []
    for i in range(1, n + 1):
        rating = rng.choices([5, 4, 3, 2, 1], weights=[62, 17, 7, 5, 9])[0]
        if rating >= 4:
            if rng.random() < 0.08:
                theme, body, sent = *_pick(rng, REVIEW_MIXED, {"quality": 1, "shipping_delivery": 1}), "Mixed"
            else:
                theme, body, sent = *_pick(rng, REVIEW_POS, REVIEW_POS_W), "Positive"
        elif rating == 3:
            if rng.random() < 0.5:
                theme, body, sent = *_pick(rng, REVIEW_MIXED, {"quality": 1, "shipping_delivery": 1}), "Mixed"
            else:
                theme, body, sent = *_pick(rng, REVIEW_NEUTRAL, {"quality": 1, "other": 1}), "Neutral"
        else:
            theme, body, sent = *_pick(rng, REVIEW_NEG, REVIEW_NEG_W), "Negative"
        body = _noise(rng, body)
        title = rng.choice(["", "", "Reseña", "Mi opinión"]) if sent != "Positive" else rng.choice(["", "", "Me encantó", "Excelente"])
        answer = f"{title}. {body}".strip(". ") if title else body
        rows.append({"title": title, "body": body, "rating": rating, "reviewer_name": f"{rng.choice(FIRST)} {rng.choice(LAST)}",
                     "reviewer_email": f"review+{i:04d}@example.com", "product_handle": rng.choice(FRAGRANCE_PRODUCTS)})
        gold.append({"role": "review", "question": "body", "answer": answer, "gold_theme": theme, "gold_sentiment": sent})
    return pd.DataFrame(rows), gold


def generate(out_dir: str | Path, seed: int = 7, gold_size: int = 400) -> dict[str, Path]:
    rng = random.Random(seed)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    paths = {}
    all_gold = []
    for name, (frame, gold) in {
        "wellness_post_purchase.csv": wellness_post_purchase(rng),
        "apparel_pdp_poll.csv": apparel_pdp(rng),
        "fragrance_reviews.csv": fragrance_reviews(rng),
    }.items():
        frame.to_csv(out / name, index=False)
        paths[name] = out / name
        for g in gold:
            g["dataset"] = name
        all_gold.extend(gold)
    sample = rng.sample(all_gold, min(gold_size, len(all_gold)))
    gold_df = pd.DataFrame(sample)
    gold_df.insert(0, "id", [f"g{i:04d}" for i in range(1, len(gold_df) + 1)])
    gold_df = gold_df[["id", "dataset", "role", "question", "answer", "gold_theme", "gold_sentiment"]]
    gold_df.to_csv(out / "gold_labels.csv", index=False)
    paths["gold_labels.csv"] = out / "gold_labels.csv"
    return paths

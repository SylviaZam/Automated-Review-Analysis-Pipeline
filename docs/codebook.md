# Codebook v2.3

Twenty-one themes plus `other`, shared by every brand so results can be compared. The themes were consolidated from manual affinity coding of post-purchase surveys, product-page polls and reviews, then extended where `voc diagnose` showed the codebook could not place a brand's answers. Each theme lists its Spanish label because stakeholder reports were delivered in Spanish.

An answer can carry several themes. The **primary** theme is the one mentioned first. `no_concern` applies only when the whole answer means "nothing" ("Ninguno", "nada en especial", "No tenía ningún miedo").

Keywords are accent-insensitive regular expressions matched at word starts. Category vocabulary lives in [../codebooks/](../codebooks/). A theme that only one brand needs belongs in a brand file passed with `--extra-themes`: those load with a `local_` prefix and are reported as brand-specific, never mixed into cross-brand comparisons.

| id | Theme | Tema | Definition | Example keywords |
|---|---|---|---|---|
| `price_value` | Price & value | Precio y valor | Price, discounts, promotions, or value for money. | precio, barat, caro, cara, costo |
| `quality` | Quality & durability | Calidad y duración | Product quality, materials, durability, longevity, or authenticity. | calidad, duracion, durabilidad, durader, dure |
| `efficacy_results` | Efficacy & results | Eficacia y resultados | Whether the product works or delivers the promised result. | funcion, sirv(a/e/en/io/ieron), resultado, efectiv, efect(o/os/iva/ivo)(?! secundario) |
| `ingredients_safety` | Ingredients & safety | Ingredientes y seguridad | What the product is made of, naturalness, side effects, allergies. | ingrediente, natural, organic, quimic, efectos? secundario |
| `trust_legitimacy` | Trust & legitimacy | Confianza en la marca | Doubts about whether the brand or site is real, safe, or trustworthy. | confia, confianza, fraude, estafa, engan |
| `shipping_delivery` | Shipping & delivery | Envío y entrega | Delivery time, shipping cost or coverage, lost or late packages. | envio, envia, entrega, lleg, paquete |
| `stock_availability` | Stock & availability | Inventario y disponibilidad | Items out of stock, restocks, or availability of a variant. | existencia, agotad, stock, disponib, volveran a tener |
| `size_fit` | Size & fit | Talla y ajuste | Sizing, measurements, fit, or dimensions. | talla, medida, tamano, horma, ajust |
| `payment_checkout` | Payment & checkout | Pago y proceso de compra | Payment methods, installments, or how to place the order. | pago, pagar, tarjeta, paypal, mercado ?pago |
| `product_info` | Product information | Información del producto | Questions about usage, dosage, benefits, or what the product is for. | como (se )?(usa/toma/aplica), (como/para que/en que) .{0,15}ayud, para que (sirve/es), modo de uso, dosis |
| `social_proof` | Recommendations & reviews | Recomendaciones y reseñas | Word of mouth, reviews, testimonials, influencers, or community. | recomend, resena, testimoni, comentarios, influencer |
| `variety_selection` | Variety & selection | Variedad de productos | Range of products, options, models, or scents to choose from. | variedad, opciones, catalogo, modelos, surtido |
| `brand_values` | Brand values | Valores de la marca | Ethics, sustainability, cruelty-free, local or Mexican-made. | cruelty, vegan, sustentab, sostenib, ecologic |
| `design_style` | Design & style | Diseño y estilo | Look, style, colors, packaging, or aesthetics. | diseno, bonit, estilo, colore?s?, moda |
| `customer_service_returns` | Service & returns | Atención y devoluciones | Customer support, responsiveness, exchanges, refunds, warranty. | atencion, servicio al cliente, devoluc, devolver, cambio de |
| `health_goal` | Health goal | Objetivo de salud | The personal health or wellbeing need the customer wants to solve. | dolor, articula, artrosis, rodilla, sop |
| `repeat_loyalty` | Repeat purchase | Recompra | Returning customers or prior positive experience with the brand. | otra vez, de nuevo, siempre (les )?compro, segunda vez, volver a comprar |
| `no_concern` | Nothing / no concern | Ninguna / sin inquietud | The customer explicitly had no concern or nothing to add. | ^(ningun[oa]?/nada/no/none/nothing/n/?a/na/ninguna duda/sin (dudas?/miedo/comentarios)/todo (bien/perfecto))$, ^no (tenia/tuve/tengo) (ningun/ninguna/miedo/dudas?), ^nothing really$, ^nada en (especial/particular)$ |
| `first_time_trial` | First-time trial | Primera compra / probar la marca | Curiosity or trying the brand for the first time, with no prior experience of it. | quiero probar, queria probar, probar la marca, probar los productos, probar el producto |
| `loyalty_rewards` | Loyalty & rewards | Lealtad y recompensas | Loyalty points, rewards, referral credit, memberships, subscriptions. | puntos, recompensa, monedero, cashback, codigo de referid |
| `site_usability` | Site & buying experience | Experiencia de compra en el sitio | How easy the store itself is to use: finding things, the size selector, the cart, errors, unclear labels or steps. | facilidad, facil de, no encuentro, no aparece, no me deja |
| `other` | Other | Otro | None of the above | |

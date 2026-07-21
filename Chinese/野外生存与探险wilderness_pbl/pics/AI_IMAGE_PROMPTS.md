# AI Image Prompts for Day 2 Camp Slides

Paste these into **ChatGPT (GPT-Image/DALL-E)** or **Midjourney** to generate
high-quality photorealistic images. Save the result to this folder using the
exact filename listed. The slide script auto-picks them up next time you run
`python3 create_day2_camp.py`.

**Style baseline for all prompts:** photorealistic landscape photography,
natural lighting, kid-friendly (no people unless noted), 16:9-ish aspect
(approximately 3:2 or 4:3 works too), high detail, no text/logos in the image.

---

## 5 Zone Slides (Good Examples)

### `zone_tent.png` — 帐篷区 (Tent Zone) — Slide 10
> A small green dome tent pitched on flat, dry, well-cleared dirt ground in a
> pine forest clearing. Bright morning light, blue sky. The ground around the
> tent is swept clean — no rocks, no sticks. A river is visible far in the
> background (clearly some distance away). Pine trees nearby but the tent is
> NOT under any tree. Photorealistic, peaceful, safe-looking campsite.
> No people. 16:9.

### `zone_fire.png` — 用火就餐区 (Fire & Eat Zone) — Slide 11
> A small campfire surrounded by a tight ring of round river stones, sitting
> on bare dirt in a forest clearing. A red bucket of water and a small pile
> of dry sand are placed beside the fire ring. A cooking pot sits on a metal
> grate over the flames. Tents are visible far in the background (clearly 4
> meters away). Late afternoon golden light. Photorealistic, organized, safe
> campsite cooking scene. No people. 16:9.

### `zone_recreation.png` — 公共娱乐区 (Recreation Area) — Slide 12
> A grassy open meadow in a sunny forest campground with a colorful soccer
> ball and a frisbee on the grass. Wildflowers around the edges. Tents are
> visible far away in the background (separate area). Bright daytime light,
> blue sky, fluffy clouds. Photorealistic, inviting, safe play area for
> kids. No people in the image, just the equipment. 16:9.

### `zone_water.png` — 取水用水区 (Water Zone) — Slide 13
> A clean mountain stream with crystal-clear water flowing over smooth round
> pebbles. A camping water filter pump and a metal pot/kettle sit on a flat
> rock at the water's edge. Two stainless-steel water bottles next to them.
> Green ferns and pine trees in the background. Soft natural light,
> photorealistic, calm and pristine. No people. 16:9.

### `zone_sanitation.png` — 卫生区 (Sanitation Zone) — Slide 14
> A simple camp washing station: a small folding table set up far from a
> campsite in a forest clearing. On the table: a bottle of hand sanitizer,
> a pack of wet wipes, a small trash bag with a closure. A trowel/small
> shovel leans against the table leg. Tall trees and undergrowth in the
> background. Soft daylight, photorealistic, organized but rustic.
> No people. 16:9.

---

## 5 "Can We Camp HERE?" Scenarios (Already Have Images — Replace for Higher Quality)

### `camp_rocks_scenario.png` — 大石头 — Slide 15
> A small green dome tent pitched at the base of a steep rocky cliff. Several
> large loose boulders are visible high above the tent, looking precarious.
> Some small stones are scattered on the ground near the tent. Cloudy,
> moody lighting suggesting unsafe conditions. Photorealistic, slightly
> ominous. The image should make a kid think "this looks dangerous."
> No people. 16:9.

### `camp grass scenario.png` — 草地 — Slide 16
> A small tent pitched in a meadow with tall, overgrown grass (knee-height)
> around it. The grass looks slightly damp, with morning dew on the tips. A
> few visible insects flying around. Sky overcast, mood is uncomfortable
> and buggy. Photorealistic. No people. 16:9.

### `camp tree scenario.png` — 大树下 — Slide 17
> A small dome tent pitched directly underneath an enormous old tree with
> visible dead branches hanging overhead. Some smaller dry sticks litter the
> ground around the tent. Dappled light filtering through the canopy. The
> setup looks unsafe — branches could fall on the tent. Photorealistic.
> No people. 16:9.

### `camp water scenario.png` — 水边 — Slide 18
> A small green tent pitched on a sandy beach, very close to ocean waves
> lapping at the shore — only a meter or two from the water. Dark stormy
> sky in the background suggesting a coming storm or high tide. Wet sand
> visible around the tent. Photorealistic, foreboding mood. No people. 16:9.

### `low or wet place scenario.png` — 低洼地 — Slide 19 (matches user's reference image!)
> A small green dome tent pitched in a low grassy depression that has
> flooded with rainwater. Standing water pools around the tent, the ground
> is saturated and muddy. Rain falling visibly, mountains in the misty
> background. Overcast gray sky. Photorealistic, somber, clearly
> uncomfortable conditions. No people. 16:9.
>
> (This is the style from your reference image — moody, photographic,
> evokes the discomfort of the choice.)

---

## How to Use

1. Copy a prompt above.
2. Paste into ChatGPT (with image generation enabled) or Midjourney/DALL-E.
3. Save the generated PNG into this `pics/` folder with the **exact filename
   shown** (e.g., `zone_tent.png`).
4. Re-run: `python3 create_day2_camp.py`
5. The slide will automatically use your image instead of the gray placeholder.

## Notes

- Aim for **landscape orientation** (wider than tall) — fits the slide's
  4.10" × 4.10" image area nicely.
- Keep file sizes reasonable (under 2 MB each ideally) so the .pptx doesn't
  bloat too much.
- If the AI keeps adding people, add `"no people, empty scene"` to the prompt.
- If you want a consistent look across all 10 images, generate them in the
  same session/style preset.

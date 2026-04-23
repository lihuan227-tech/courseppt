# College Application Platforms PPT — Design Spec

## Overview

A 18-slide `.pptx` teaching presentation for 8-10 high school students covering Common App and UC application systems. 90-minute class session. English only. Generated via `python-pptx`.

**Approach:** Compare-First, Detail-Later — introduce the 3E framework (Education, Experience, Exposure) through guided discovery, then deep-dive into each platform, compare them, and end with self-reflection and planning.

## Visual Design

- **Color scheme:** Teal/green matching existing GR EDU branding (primary teal `#00897b`, accent green `#9DC64D`, dark text `#2C2C2C`)
- **Font:** KaiTi for consistency with existing PPT generators
- **Slide size:** Widescreen 16:9 (13.333" x 7.5") — standard presentation format
- **Branding:** GR EDU logo top-left on each slide
- **Style reference:** Existing 3E slide (teal cards with icons, illustrated characters)

## Slide-by-Slide Structure

### Section 1: Opening (Slides 1-4)

**Slide 1 — Title Slide**
- Title: "College Application Platforms: What You Need to Know"
- Subtitle: "Common App vs. UC Application"
- GR EDU branding

**Slide 2 — Overview: Application Platforms**
- Visual list of major platforms:
  - Common App (900+ schools)
  - UC Application (9 UC campuses)
  - Coalition App / SCOIR
  - ApplyTexas, MIT, Georgetown (self-hosted)
- Callout box: "Today's Focus: Common App & UC"

**Slide 3 — What Are They Asking You?**
- Side-by-side: Common App sections vs. UC application sections
- Common App: Demographics, Education, Testing, Activities (10 slots), Essay (650 words), Recommendations
- UC: Personal Info, Campuses, Academic History, Activities (20 slots), PIQs (4 x 350 words)
- Discussion prompt: "Look at these questions — what do you think colleges are trying to learn about you?"
- 3-5 minute student discussion

**Slide 4 — The 3E Framework (Reveal)**
- Teal card design (matching existing 3E slide):
  - Education: GPA, Course Rigor, Standardized Tests, Recommendations
  - Experience: Activities, Leadership, Community Service, Work/Internships
  - Exposure: Competitions & Awards, Publications, Research, Public Impact
- Activity: Students categorize application sections from Slide 3 into 3E columns on worksheet/notes

### Section 2: Common App Deep Dive (Slides 5-8)

**Slide 5 — Common App: Activities Section**
- 10 activity slots, 150 characters per description
- Activity categories list
- "150 characters = ~2 sentences. Every word counts."
- Good vs. bad activity description example side by side

**Slide 6 — Common App: Essay (Personal Statement)**
- 7 prompt options, 250-650 words
- All 7 prompts listed briefly
- Highlight: Prompt 7 is open-ended
- "This is your ONE chance to show personality beyond grades and scores"

**Slide 7 — Common App: Supplemental Essays**
- Many schools add their own essays through Common App
- Common types: "Why this school?", "Why this major?", "Community/diversity", "Extracurricular elaboration"
- Tip: "Research each school — generic answers are obvious"

**Slide 8 — Common App: Other Sections**
- Demographics, Family, Education, Testing, Honors (5 slots)
- Straightforward but strategic (e.g., which test scores to report)

### Section 3: UC Application Deep Dive (Slides 9-12)

**Slide 9 — UC: Activities & Awards**
- 20 activity slots (vs. Common App's 10)
- Separate Awards/Honors section (5 slots)
- Categories: Volunteer, Work, Awards, Educational Prep Programs
- "Start tracking your activities NOW."

**Slide 10 — UC: Personal Insight Questions (PIQs)**
- 8 prompts, choose 4, 350 words max each
- All 8 PIQ prompts listed
- Key difference: shorter, more focused, 4 essays instead of 1
- Strategy: "Pick 4 that showcase different sides of you"

**Slide 11 — UC: What Makes It Different**
- No letters of recommendation (most applicants)
- No supplemental essays — PIQs serve that role
- UC GPA calculated differently (honors/AP bonus, 10th-11th weighted)
- Same application for all 9 UC campuses

**Slide 12 — UC: Holistic Review — The 13 Criteria**
- UC's published 13 review criteria listed
- Connection back to 3E framework

### Section 4: Comparison & Planning (Slides 13-15)

**Slide 13 — Side-by-Side Comparison Chart**

| | Common App | UC Application |
|---|---|---|
| Schools | 900+ nationwide | 9 UC campuses |
| Activities | 10 slots, 150 chars | 20 slots + 5 awards |
| Essays | 1 personal (650 words) + supplementals | 4 PIQs (350 words each) |
| Recommendations | 1-3 required by most schools | Not required |
| Test Scores | School-dependent | Test-optional |
| GPA | School-reported | UC-calculated (10th-11th) |
| Deadline | Varies (Nov-Jan typical) | Nov 30 |

- Discussion: "What should you prepare differently for each?"

**Slide 14 — The Timeline: When to Start What**
- Visual timeline: 9th > 10th > 11th > 12th
  - 9th-10th: Build activities, explore interests, keep grades up
  - Summer before 11th: Start leadership roles, competitions, research
  - 11th: SAT/ACT, build activity list, begin essay brainstorming
  - Summer before 12th: Draft essays, finalize activity descriptions
  - 12th fall: Submit applications
- "You can't cram a 4-year story in senior fall"

**Slide 15 — 3E Self-Audit (Worksheet Activity)**
- 3-column table: Education / Experience / Exposure
- Row 1: What I currently have
- Row 2: Gaps — what's missing?
- Row 3: Action items — what can I do this summer/next year?
- 10-minute activity

### Section 5: Reflection & Closing (Slides 16-18)

**Slide 16 — Reflection: Where Am I Now?**
- "Which 3E category is your strongest? Which has the biggest gap?"
- "List 3 activities you could put on an application today."
- "If you had to write your Common App essay now, which prompt would you pick? Why?"

**Slide 17 — Reflection: What's My Plan?**
- "Name one activity you can start or deepen this summer."
- "What can you do this month to strengthen your weakest 3E area?"
- "Who could write you a strong recommendation letter? What do they know about you?"
- "Set a goal: By [date], I will have ___________."

**Slide 18 — Key Takeaways & Next Steps**
- Know the platforms — Common App and UC have different requirements
- Essays and activities need years of preparation, not weeks
- Use the 3E framework to audit yourself regularly
- Start now — the best application tells a story built over time
- GR EDU contact / next class info
- "Questions?"

## Technical Implementation

- Generated via `python-pptx` script at `English/create_college_app_ppt.py`
- Output: `English/college_app_platforms.pptx`
- Slide size: 16:9 widescreen
- Helper functions: reuse patterns from existing `create_flyer.py` scripts (add_shape, set_text, add_para)
- No external images required — use shapes and text styling for visual design
- Comparison table on Slide 13 built with pptx table object
- Timeline on Slide 14 built with shapes (rounded rectangles + connector lines)
- 3E cards on Slide 4 use teal rounded rectangles matching the reference screenshot style

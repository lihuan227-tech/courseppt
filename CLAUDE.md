# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Summer course website and marketing materials for **谷雨中文 GR EDU** — a Chinese education center offering summer elective classes in Chinese, English, and Math for K–5th grade students. All content is bilingual (Chinese/English).

The HTML pages are embedded into **Google Sites** via iframe. They must be fully self-contained (inline CSS, no external JS frameworks except Google Fonts and html2canvas CDN).

## Architecture

- **`summer schedule/`** — Core scheduling, registration, and tuition pages (embedded in Google Sites)
  - `master_schedule.html` — Interactive **summer** master schedule with course selector, weekly grid, conflict detection, tuition calculator, print/save-image functionality
  - `spring_schedule.html` — Interactive **spring after-school** schedule (Cupertino site, 7 subjects: Chinese/English/Math/Art/Chess/Speech/Pingpong). Uses `data-slots` attribute with `Day:startMin-endMin` format to handle classes with different times on different days.
  - `elective_tuition.html` — Tuition pricing table
  - `course_selector.html` — Standalone course selection tool
  - `schedule_chinese.html`, `schedule_english.html`, `schedule_math.html` — Per-subject schedule tables
  - `registration_form_template.html` — Visual reference template with interactive calendar
  - `create_google_form.gs` — Google Apps Script that auto-creates the Google Form with onSubmit trigger for tuition calculation and confirmation emails
- **`Chinese/`**, **`English/`**, **`Math/`** — Per-subject course detail pages and flyer generators
  - `create_flyer*.py` — Python scripts using `python-pptx` to generate editable PPT flyers

## Design System

**Subject colors** (used consistently across all files):
- Chinese: `#FF8C00` (orange), headers `#f59d00`
- English: `#1976d2` (blue), headers `#1e88e5`
- Math: `#388e3c` (green), headers `#43a047`
- Art: `#9c27b0` (purple), headers `#ab47bc` — spring only
- Chess: `#795548` (brown), headers `#8d6e63` — spring only
- Speech & Debate: `#e91e63` (pink), headers `#ec407a` — spring only
- Pingpong: `#00897b` (teal), headers `#26a69a` — spring only

**Fonts:** `'Noto Sans SC', 'Kaiti SC', 'KaiTi', sans-serif` — loaded via Google Fonts CDN

**Layout:** `max-width: 1100px` for all embeddable pages

## Key Constraints

- **Google Sites iframe restrictions**: `window.print()` is blocked in cross-origin iframes. Print uses a 3-level fallback: `window.open` → `top.open` → save-as-image via html2canvas. Mobile uses native `navigator.share` when available.
- **No build system**: HTML files are standalone. Open directly in browser or embed in Google Sites.
- **PPT generation**: `pip install python-pptx Pillow` then `python create_flyer.py`

## Session/Date Structure

4 sessions across 8 weeks (no class week of 6/30):
- Session 1: Week 1 (6/8–6/12), Week 2 (6/15–6/19)
- Session 2: Week 3 (6/22–6/26), Week 4 (7/6–7/10)
- Session 3: Week 5 (7/13–7/17), Week 6 (7/20–7/24)
- Session 4: Week 7 (7/27–7/31), Week 8 (8/3–8/7)

## Tuition

- Chinese: $120/session
- English: $160/session
- Math: $160/session
- Payment: Zelle to gredu2019@gmail.com

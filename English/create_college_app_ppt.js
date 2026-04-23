const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const { FaGraduationCap, FaUniversity, FaPencilAlt, FaListUl, FaBalanceScale,
        FaClock, FaClipboardCheck, FaQuestionCircle, FaLightbulb, FaRocket,
        FaChalkboardTeacher, FaBullseye, FaUsers, FaTrophy, FaBookOpen,
        FaStar, FaCalendarAlt, FaArrowRight, FaCheckCircle, FaSearch } = require("react-icons/fa");
const { MdSchool, MdWork, MdPublic, MdCompareArrows } = require("react-icons/md");

// ── COLORS ──
const C = {
  teal:       "00897B",
  tealDark:   "00695C",
  tealLight:  "E0F2F1",
  green:      "43A047",
  greenLight: "E8F5E9",
  accent:     "FF8C00",
  accentDark: "E67E00",
  dark:       "1A1A1A",
  text:       "2C2C2C",
  textMid:    "555555",
  textLight:  "757575",
  white:      "FFFFFF",
  offWhite:   "F5F7FA",
  cream:      "FAFBFC",
  cardBg:     "F8FFFE",
  border:     "B2DFDB",
  tableTeal:  "00897B",
  tableAlt:   "F1F8F7",
};

// ── ICON HELPER ──
function renderIconSvg(IconComponent, color, size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color: "#" + color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// ── FACTORY HELPERS (avoid mutable reuse) ──
const cardShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 });
const softShadow = () => ({ type: "outer", blur: 6, offset: 3, angle: 135, color: "000000", opacity: 0.12 });

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10" x 5.625"
  pres.author = "GR EDU";
  pres.title = "College Application Platforms: Common App vs UC";

  // Pre-render all icons
  const icons = {};
  const iconMap = {
    grad: [FaGraduationCap, C.teal],
    uni: [FaUniversity, C.teal],
    pencil: [FaPencilAlt, C.white],
    list: [FaListUl, C.white],
    balance: [FaBalanceScale, C.teal],
    clock: [FaClock, C.accent],
    clipboard: [FaClipboardCheck, C.teal],
    question: [FaQuestionCircle, C.accent],
    lightbulb: [FaLightbulb, C.accent],
    rocket: [FaRocket, C.teal],
    teacher: [FaChalkboardTeacher, C.teal],
    bullseye: [FaBullseye, C.teal],
    users: [FaUsers, C.white],
    trophy: [FaTrophy, C.white],
    book: [FaBookOpen, C.white],
    star: [FaStar, C.accent],
    calendar: [FaCalendarAlt, C.teal],
    arrow: [FaArrowRight, C.white],
    check: [FaCheckCircle, C.teal],
    search: [FaSearch, C.teal],
    school: [MdSchool, C.white],
    work: [MdWork, C.white],
    publicIcon: [MdPublic, C.white],
    compare: [MdCompareArrows, C.teal],
    // 3E category icons (white on teal)
    edu: [FaGraduationCap, C.white],
    exp: [MdWork, C.white],
    expo: [FaTrophy, C.white],
    // Additional colored versions
    gradWhite: [FaGraduationCap, C.white],
    uniWhite: [FaUniversity, C.white],
    rocketWhite: [FaRocket, C.white],
    checkWhite: [FaCheckCircle, C.white],
    lightbulbWhite: [FaLightbulb, C.white],
    starWhite: [FaStar, C.white],
    calendarWhite: [FaCalendarAlt, C.white],
  };
  for (const [key, [Comp, color]] of Object.entries(iconMap)) {
    icons[key] = await iconToBase64Png(Comp, color);
  }

  // ════════════════════════════════════════════════
  // SLIDE 1 — Title
  // ════════════════════════════════════════════════
  let slide = pres.addSlide();
  slide.background = { color: C.tealDark };
  // Decorative shapes
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });
  // Large icon
  slide.addImage({ data: icons.gradWhite, x: 4.4, y: 0.5, w: 1.2, h: 1.2 });
  // Title
  slide.addText("College Application Platforms", {
    x: 0.5, y: 1.8, w: 9, h: 1.0,
    fontSize: 40, fontFace: "Georgia", bold: true,
    color: C.white, align: "center", margin: 0
  });
  slide.addText("What You Need to Know", {
    x: 0.5, y: 2.7, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", italic: true,
    color: C.accent, align: "center", margin: 0
  });
  // Divider
  slide.addShape(pres.shapes.LINE, {
    x: 3.5, y: 3.6, w: 3, h: 0, line: { color: C.border, width: 1.5 }
  });
  slide.addText("Common App  vs.  UC Application", {
    x: 0.5, y: 3.8, w: 9, h: 0.6,
    fontSize: 22, fontFace: "Calibri",
    color: C.tealLight, align: "center", margin: 0
  });
  slide.addText("GR EDU  |  谷雨中文", {
    x: 0.5, y: 4.7, w: 9, h: 0.5,
    fontSize: 14, fontFace: "Calibri",
    color: C.border, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 2 — Overview: Application Platforms
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("Application Platforms Overview", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Platform cards - 2 rows of 3
  const platforms = [
    { name: "Common App", desc: "900+ schools nationwide", icon: icons.uni, highlight: true },
    { name: "UC Application", desc: "9 UC campuses", icon: icons.grad, highlight: true },
    { name: "Coalition / SCOIR", desc: "150+ member schools", icon: icons.users, highlight: false },
    { name: "ApplyTexas", desc: "Texas public universities", icon: icons.school, highlight: false },
    { name: "MIT / Georgetown", desc: "Self-hosted portals", icon: icons.book, highlight: false },
    { name: "CSS Profile", desc: "Financial aid (400+ schools)", icon: icons.clipboard, highlight: false },
  ];

  for (let i = 0; i < 6; i++) {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 3.1;
    const y = 1.2 + row * 2.0;
    const p = platforms[i];
    const bg = p.highlight ? C.teal : C.white;
    const txtColor = p.highlight ? C.white : C.text;
    const descColor = p.highlight ? C.tealLight : C.textMid;

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.8, h: 1.7,
      fill: { color: bg },
      shadow: cardShadow(),
      rectRadius: 0.05
    });
    // Icon circle
    if (p.highlight) {
      slide.addImage({ data: p.icon, x: x + 1.05, y: y + 0.15, w: 0.55, h: 0.55 });
    } else {
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.95, y: y + 0.1, w: 0.7, h: 0.7,
        fill: { color: C.teal }
      });
      slide.addImage({ data: p.icon, x: x + 1.05, y: y + 0.2, w: 0.5, h: 0.5 });
    }
    slide.addText(p.name, {
      x: x + 0.15, y: y + 0.8, w: 2.5, h: 0.4,
      fontSize: 15, fontFace: "Calibri", bold: true,
      color: txtColor, align: "center", margin: 0
    });
    slide.addText(p.desc, {
      x: x + 0.15, y: y + 1.15, w: 2.5, h: 0.4,
      fontSize: 12, fontFace: "Calibri",
      color: descColor, align: "center", margin: 0
    });
  }

  // "Today's Focus" callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 2.5, y: 5.05, w: 5, h: 0.45,
    fill: { color: C.accent }
  });
  slide.addText("Today's Focus:  Common App  &  UC Application", {
    x: 2.5, y: 5.05, w: 5, h: 0.45,
    fontSize: 15, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 3 — What Are They Asking You?
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addImage({ data: icons.searchWhite || icons.lightbulbWhite, x: 0.4, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("What Are They Asking You?", {
    x: 1.0, y: 0.1, w: 8.5, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Common App column
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.3, h: 3.5,
    fill: { color: C.white }, shadow: cardShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.3, h: 0.5,
    fill: { color: C.teal }
  });
  slide.addText("Common App Sections", {
    x: 0.5, y: 1.1, w: 4.1, h: 0.5,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });
  slide.addText([
    { text: "Profile & Demographics", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Family Information", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Education History", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Testing (SAT/ACT/AP)", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Activities — 10 slots, 150 chars each", options: { bullet: true, breakLine: true, fontSize: 13, bold: true, color: C.accent } },
    { text: "Honors — 5 slots", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Personal Essay — 650 words (1 essay)", options: { bullet: true, breakLine: true, fontSize: 13, bold: true, color: C.accent } },
    { text: "Supplemental Essays (school-specific)", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Recommendations (1-3 letters)", options: { bullet: true, fontSize: 13 } },
  ], {
    x: 0.6, y: 1.7, w: 3.9, h: 2.8,
    fontFace: "Calibri", color: C.text, paraSpaceAfter: 4, valign: "top"
  });

  // UC column
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.1, w: 4.3, h: 3.5,
    fill: { color: C.white }, shadow: cardShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.1, w: 4.3, h: 0.5,
    fill: { color: C.tealDark }
  });
  slide.addText("UC Application Sections", {
    x: 5.3, y: 1.1, w: 4.1, h: 0.5,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });
  slide.addText([
    { text: "Personal Information", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Campus & Major Selection", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Academic History", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Test Scores (optional)", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Activities & Awards — 20 + 5 slots", options: { bullet: true, breakLine: true, fontSize: 13, bold: true, color: C.accent } },
    { text: "Scholarships & Programs", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Personal Insight Questions — 4 × 350 words", options: { bullet: true, breakLine: true, fontSize: 13, bold: true, color: C.accent } },
    { text: "No recommendation letters required", options: { bullet: true, breakLine: true, fontSize: 13 } },
    { text: "Additional Comments (optional)", options: { bullet: true, fontSize: 13 } },
  ], {
    x: 5.4, y: 1.7, w: 3.9, h: 2.8,
    fontFace: "Calibri", color: C.text, paraSpaceAfter: 4, valign: "top"
  });

  // Discussion prompt
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.8, w: 9.2, h: 0.65,
    fill: { color: C.accent }
  });
  slide.addImage({ data: icons.lightbulbWhite, x: 0.6, y: 4.87, w: 0.4, h: 0.4 });
  slide.addText("Discussion: Look at these questions — what do you think colleges are trying to learn about you?", {
    x: 1.1, y: 4.8, w: 8.3, h: 0.65,
    fontSize: 15, fontFace: "Calibri", bold: true, italic: true,
    color: C.white, valign: "middle", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 4 — The 3E Framework (Reveal)
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("3E STRUCTURE", {
    x: 0.5, y: 0.05, w: 4, h: 0.35,
    fontSize: 14, fontFace: "Calibri", bold: true, charSpacing: 4,
    color: C.accent, margin: 0
  });
  slide.addText("What Colleges Are Looking For in Strong Candidates", {
    x: 0.5, y: 0.35, w: 9, h: 0.5,
    fontSize: 22, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // 3E Cards
  const threeE = [
    { label: "Education", icon: icons.edu, items: ["GPA", "Course Rigor", "Standardized Tests", "Teacher & Counselor Recommendations"] },
    { label: "Experience", icon: icons.exp, items: ["Extracurricular Activities", "Leadership", "Community Service", "Work, Internships, Family Responsibilities"] },
    { label: "Exposure", icon: icons.expo, items: ["Competitions & Awards", "Publications & Research", "Projects with Public Impact", "Awards from City/County/State"] },
  ];

  for (let i = 0; i < 3; i++) {
    const y = 1.15 + i * 1.45;
    const e = threeE[i];
    // Teal label card
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 1.8, h: 1.2,
      fill: { color: C.teal }, shadow: cardShadow()
    });
    slide.addImage({ data: e.icon, x: 1.05, y: y + 0.1, w: 0.6, h: 0.6 });
    slide.addText(e.label, {
      x: 0.5, y: y + 0.7, w: 1.8, h: 0.4,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.white, align: "center", margin: 0
    });
    // Bullet items
    const bullets = e.items.map((item, idx) => ({
      text: item,
      options: { bullet: true, fontSize: 14, breakLine: idx < e.items.length - 1 }
    }));
    slide.addText(bullets, {
      x: 2.6, y, w: 4.0, h: 1.2,
      fontFace: "Calibri", color: C.text, valign: "middle", paraSpaceAfter: 3
    });
  }

  // Activity prompt on right side
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.8, y: 1.15, w: 2.8, h: 4.2,
    fill: { color: C.tealLight }, shadow: cardShadow()
  });
  slide.addImage({ data: icons.pencil, x: 7.0, y: 1.15, w: 0.01, h: 0.01 }); // skip white pencil
  slide.addText("ACTIVITY", {
    x: 6.9, y: 1.3, w: 2.6, h: 0.35,
    fontSize: 13, fontFace: "Calibri", bold: true, charSpacing: 3,
    color: C.tealDark, align: "center", margin: 0
  });
  slide.addText([
    { text: "Categorize the application sections from Slide 3:", options: { breakLine: true, fontSize: 13, bold: true } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "Which parts are...", options: { breakLine: true, fontSize: 12, italic: true } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "Education?", options: { breakLine: true, fontSize: 14, bold: true, color: C.teal } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "Experience?", options: { breakLine: true, fontSize: 14, bold: true, color: C.teal } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "Exposure?", options: { breakLine: true, fontSize: 14, bold: true, color: C.teal } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "Fill in your worksheet / take notes!", options: { fontSize: 12, italic: true, color: C.accent } },
  ], {
    x: 7.0, y: 1.7, w: 2.4, h: 3.4,
    fontFace: "Calibri", color: C.text, valign: "top"
  });

  // ════════════════════════════════════════════════
  // SLIDE 5 — Common App: Activities Section
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("Common App: Activities Section", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Key stats
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 2.8, h: 1.6,
    fill: { color: C.teal }, shadow: cardShadow()
  });
  slide.addText("10", {
    x: 0.4, y: 1.15, w: 2.8, h: 0.8,
    fontSize: 52, fontFace: "Georgia", bold: true,
    color: C.white, align: "center", margin: 0
  });
  slide.addText("Activity Slots", {
    x: 0.4, y: 1.9, w: 2.8, h: 0.35,
    fontSize: 16, fontFace: "Calibri",
    color: C.tealLight, align: "center", margin: 0
  });
  slide.addText("150 chars each", {
    x: 0.4, y: 2.2, w: 2.8, h: 0.35,
    fontSize: 13, fontFace: "Calibri",
    color: C.border, align: "center", margin: 0
  });

  // Categories
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 3.5, y: 1.1, w: 6.1, h: 1.6,
    fill: { color: C.white }, shadow: cardShadow()
  });
  slide.addText("Activity Categories", {
    x: 3.7, y: 1.15, w: 5.7, h: 0.35,
    fontSize: 15, fontFace: "Calibri", bold: true,
    color: C.teal, margin: 0
  });
  const categories = ["Athletics", "Community Service", "Music / Art", "Academic Clubs",
    "Work (Paid)", "Family Responsibilities", "Internship", "Research",
    "Student Gov't", "Volunteering"];
  // 2 columns of tags
  for (let i = 0; i < categories.length; i++) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 3.7 + col * 2.9, y: 1.55 + row * 0.22, w: 2.6, h: 0.2,
      fill: { color: C.tealLight }
    });
    slide.addText(categories[i], {
      x: 3.7 + col * 2.9, y: 1.53 + row * 0.22, w: 2.6, h: 0.22,
      fontSize: 11, fontFace: "Calibri", color: C.tealDark, align: "center", margin: 0
    });
  }

  // Character count callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 2.9, w: 9.2, h: 0.5,
    fill: { color: C.accent }
  });
  slide.addText("150 characters = ~2 sentences.  Every. Word. Counts.", {
    x: 0.4, y: 2.9, w: 9.2, h: 0.5,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });

  // Good vs Bad example
  slide.addText("Example: Describing Your Activity", {
    x: 0.5, y: 3.6, w: 9, h: 0.35,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.teal, margin: 0
  });

  // Bad example
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.0, w: 4.4, h: 1.3,
    fill: { color: "FFF0F0" }, shadow: cardShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.0, w: 4.4, h: 0.35,
    fill: { color: "D32F2F" }
  });
  slide.addText("Weak Description", {
    x: 0.5, y: 4.0, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Calibri", bold: true,
    color: C.white, margin: 0
  });
  slide.addText("I was in the math club and we did competitions and I helped tutor younger students sometimes.", {
    x: 0.6, y: 4.4, w: 4.0, h: 0.8,
    fontSize: 12, fontFace: "Calibri", italic: true,
    color: C.textMid, margin: 0
  });

  // Good example
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 4.0, w: 4.4, h: 1.3,
    fill: { color: C.greenLight }, shadow: cardShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 4.0, w: 4.4, h: 0.35,
    fill: { color: C.green }
  });
  slide.addText("Strong Description", {
    x: 5.3, y: 4.0, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: "Calibri", bold: true,
    color: C.white, margin: 0
  });
  slide.addText("Founded Math Olympiad prep program; coached 15 students weekly, 4 qualified for MATHCOUNTS States. Led curriculum design.", {
    x: 5.4, y: 4.4, w: 4.0, h: 0.8,
    fontSize: 12, fontFace: "Calibri", italic: true,
    color: C.text, margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 6 — Common App: Personal Essay
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("Common App: Personal Essay", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Stats bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.2, h: 0.6,
    fill: { color: C.teal }
  });
  slide.addText("7 Prompts  |  250-650 Words  |  1 Essay", {
    x: 0.4, y: 1.1, w: 4.2, h: 0.6,
    fontSize: 15, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });

  // 7 prompts listed
  const prompts = [
    "1. Background, identity, interest, or talent that is meaningful to you",
    "2. A lesson from an obstacle, setback, or failure",
    "3. A time you questioned or challenged a belief or idea",
    "4. A problem you'd like to solve — or have solved",
    "5. Personal growth or a new understanding of yourself",
    "6. A topic, idea, or concept that captivates you",
    "7. Share an essay on any topic of your choice",
  ];
  const promptBullets = prompts.map((p, i) => ({
    text: p,
    options: { fontSize: 12, breakLine: i < prompts.length - 1, color: i === 6 ? C.accent : C.text, bold: i === 6 }
  }));
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.8, w: 5.6, h: 3.5,
    fill: { color: C.white }, shadow: cardShadow()
  });
  slide.addText("Essay Prompt Options", {
    x: 0.6, y: 1.85, w: 5.2, h: 0.35,
    fontSize: 14, fontFace: "Calibri", bold: true,
    color: C.teal, margin: 0
  });
  slide.addText(promptBullets, {
    x: 0.6, y: 2.2, w: 5.2, h: 3.0,
    fontFace: "Calibri", color: C.text, valign: "top", paraSpaceAfter: 5
  });

  // Key insight box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: 1.1, w: 3.3, h: 4.2,
    fill: { color: C.tealLight }, shadow: cardShadow()
  });
  slide.addImage({ data: icons.star, x: 7.55, y: 1.3, w: 0.6, h: 0.6 });
  slide.addText("Key Insight", {
    x: 6.4, y: 1.95, w: 3.1, h: 0.35,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.tealDark, align: "center", margin: 0
  });
  slide.addText([
    { text: "This is your ONE chance to show personality beyond grades and scores.", options: { breakLine: true, fontSize: 14, bold: true } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "Admissions officers read thousands of essays.", options: { breakLine: true, fontSize: 13 } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "Be specific.", options: { breakLine: true, fontSize: 14, bold: true, color: C.accent } },
    { text: "Be authentic.", options: { breakLine: true, fontSize: 14, bold: true, color: C.accent } },
    { text: "Be memorable.", options: { fontSize: 14, bold: true, color: C.accent } },
  ], {
    x: 6.5, y: 2.4, w: 2.9, h: 2.7,
    fontFace: "Calibri", color: C.text, valign: "top"
  });

  // ════════════════════════════════════════════════
  // SLIDE 7 — Common App: Supplemental Essays
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("Common App: Supplemental Essays", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  slide.addText("Many schools require additional essays through Common App", {
    x: 0.5, y: 1.1, w: 9, h: 0.4,
    fontSize: 15, fontFace: "Calibri", italic: true,
    color: C.textMid, margin: 0
  });

  // 4 common types as cards
  const suppTypes = [
    { title: "\"Why This School?\"", desc: "Show you've researched the school. Mention specific programs, professors, clubs, or values that draw you." },
    { title: "\"Why This Major?\"", desc: "Connect your interests, experiences, and goals. Show a genuine intellectual curiosity, not just career plans." },
    { title: "\"Community & Diversity\"", desc: "How will you contribute to campus? What perspective or experience do you bring that others might not?" },
    { title: "\"Elaborate on Activity\"", desc: "Go deeper on one extracurricular. This is your chance to show passion and impact beyond 150 characters." },
  ];
  for (let i = 0; i < 4; i++) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.7;
    const y = 1.7 + row * 1.7;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.4, h: 1.45,
      fill: { color: C.white }, shadow: cardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: 1.45,
      fill: { color: C.teal }
    });
    slide.addText(suppTypes[i].title, {
      x: x + 0.2, y: y + 0.08, w: 4.0, h: 0.35,
      fontSize: 15, fontFace: "Calibri", bold: true,
      color: C.teal, margin: 0
    });
    slide.addText(suppTypes[i].desc, {
      x: x + 0.2, y: y + 0.45, w: 4.0, h: 0.9,
      fontSize: 12, fontFace: "Calibri",
      color: C.text, margin: 0
    });
  }

  // Tip
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 5.05, w: 9.2, h: 0.45,
    fill: { color: C.accent }
  });
  slide.addText("Tip: Research each school — generic answers are obvious to admissions officers.", {
    x: 0.4, y: 5.05, w: 9.2, h: 0.45,
    fontSize: 14, fontFace: "Calibri", bold: true, italic: true,
    color: C.white, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 8 — Common App: Other Sections
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("Common App: Other Sections", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const otherSections = [
    { title: "Demographics & Family", desc: "Straightforward personal info. Be accurate — schools verify.", icon: icons.users },
    { title: "Education", desc: "School info, GPA, class rank. Report what your school provides.", icon: icons.school },
    { title: "Testing", desc: "SAT, ACT, AP, IB scores. Strategic: only send scores that help you.", icon: icons.clipboard },
    { title: "Honors (5 slots)", desc: "Academic distinctions: Honor Roll, AP Scholar, National Merit, subject awards.", icon: icons.trophy },
  ];
  for (let i = 0; i < 4; i++) {
    const y = 1.2 + i * 1.05;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.85,
      fill: { color: i % 2 === 0 ? C.white : C.tealLight }, shadow: cardShadow()
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y + 0.12, w: 0.6, h: 0.6,
      fill: { color: C.teal }
    });
    slide.addImage({ data: otherSections[i].icon, x: 0.78, y: y + 0.2, w: 0.4, h: 0.4 });
    slide.addText(otherSections[i].title, {
      x: 1.5, y: y + 0.05, w: 7.8, h: 0.35,
      fontSize: 16, fontFace: "Calibri", bold: true,
      color: C.teal, margin: 0
    });
    slide.addText(otherSections[i].desc, {
      x: 1.5, y: y + 0.4, w: 7.8, h: 0.35,
      fontSize: 13, fontFace: "Calibri",
      color: C.text, margin: 0
    });
  }

  // Strategic note
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.0, w: 9, h: 0.45,
    fill: { color: C.tealLight }
  });
  slide.addText("These sections seem simple, but strategic choices (which scores to send, how to frame honors) matter.", {
    x: 0.7, y: 5.0, w: 8.6, h: 0.45,
    fontSize: 13, fontFace: "Calibri", italic: true,
    color: C.tealDark, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 9 — UC: Activities & Awards
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.tealDark } });
  slide.addText("UC Application: Activities & Awards", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Two big stat cards
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.1, w: 4.3, h: 1.5,
    fill: { color: C.tealDark }, shadow: cardShadow()
  });
  slide.addText("20", {
    x: 0.4, y: 1.15, w: 4.3, h: 0.8,
    fontSize: 56, fontFace: "Georgia", bold: true,
    color: C.white, align: "center", margin: 0
  });
  slide.addText("Activity Slots", {
    x: 0.4, y: 1.9, w: 4.3, h: 0.3,
    fontSize: 16, fontFace: "Calibri",
    color: C.tealLight, align: "center", margin: 0
  });
  slide.addText("vs. Common App's 10", {
    x: 0.4, y: 2.2, w: 4.3, h: 0.25,
    fontSize: 12, fontFace: "Calibri", italic: true,
    color: C.border, align: "center", margin: 0
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.1, w: 4.3, h: 1.5,
    fill: { color: C.accent }, shadow: cardShadow()
  });
  slide.addText("5", {
    x: 5.3, y: 1.15, w: 4.3, h: 0.8,
    fontSize: 56, fontFace: "Georgia", bold: true,
    color: C.white, align: "center", margin: 0
  });
  slide.addText("Awards / Honors Slots", {
    x: 5.3, y: 1.9, w: 4.3, h: 0.3,
    fontSize: 16, fontFace: "Calibri",
    color: C.white, align: "center", margin: 0
  });
  slide.addText("Separate section", {
    x: 5.3, y: 2.2, w: 4.3, h: 0.25,
    fontSize: 12, fontFace: "Calibri", italic: true,
    color: C.white, align: "center", margin: 0
  });

  // UC categories
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 2.9, w: 9.2, h: 2.3,
    fill: { color: C.white }, shadow: cardShadow()
  });
  slide.addText("UC Activity Categories", {
    x: 0.6, y: 2.95, w: 8.8, h: 0.35,
    fontSize: 15, fontFace: "Calibri", bold: true,
    color: C.teal, margin: 0
  });
  const ucCats = [
    "Award or Honor", "Educational Prep Program", "Extracurricular Activity",
    "Other Coursework", "Volunteer / Community Service", "Work Experience"
  ];
  for (let i = 0; i < ucCats.length; i++) {
    const col = i % 3;
    const row = Math.floor(i / 3);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.7 + col * 3.0, y: 3.4 + row * 0.8, w: 2.7, h: 0.55,
      fill: { color: C.tealLight }
    });
    slide.addText(ucCats[i], {
      x: 0.7 + col * 3.0, y: 3.4 + row * 0.8, w: 2.7, h: 0.55,
      fontSize: 13, fontFace: "Calibri", bold: true,
      color: C.tealDark, align: "center", valign: "middle", margin: 0
    });
  }

  // Callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 5.05, w: 9.2, h: 0.45,
    fill: { color: C.accent }
  });
  slide.addText("More space = more to fill. Start tracking your activities NOW.", {
    x: 0.4, y: 5.05, w: 9.2, h: 0.45,
    fontSize: 15, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 10 — UC: PIQs
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.tealDark } });
  slide.addText("UC: Personal Insight Questions (PIQs)", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 26, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Stats
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 0.55,
    fill: { color: C.teal }
  });
  slide.addText("8 Prompts  |  Choose 4  |  350 Words Max Each", {
    x: 0.4, y: 1.05, w: 9.2, h: 0.55,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });

  // 8 PIQ prompts in 2 columns
  const piqs = [
    "1. Leadership experience",
    "2. Creative side — approach to problem-solving",
    "3. Greatest talent or skill",
    "4. Educational opportunity or barrier you've overcome",
    "5. Significant challenge — what did you do?",
    "6. Favorite academic subject — what makes it exciting?",
    "7. Community contribution",
    "8. What sets you apart?",
  ];
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.8, w: 9.2, h: 2.6,
    fill: { color: C.white }, shadow: cardShadow()
  });
  for (let i = 0; i < 8; i++) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    slide.addText(piqs[i], {
      x: 0.6 + col * 4.5, y: 1.9 + row * 0.6, w: 4.2, h: 0.5,
      fontSize: 13, fontFace: "Calibri",
      color: C.text, valign: "middle", margin: 0
    });
  }

  // Strategy box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.55, w: 9.2, h: 0.9,
    fill: { color: C.tealLight }
  });
  slide.addImage({ data: icons.lightbulb, x: 0.6, y: 4.65, w: 0.5, h: 0.5 });
  slide.addText([
    { text: "Strategy: ", options: { bold: true, fontSize: 14, color: C.tealDark } },
    { text: "Pick 4 that showcase ", options: { fontSize: 14 } },
    { text: "different ", options: { fontSize: 14, bold: true, italic: true, color: C.accent } },
    { text: "sides of you. Don't repeat the same story across multiple PIQs.", options: { fontSize: 14 } },
  ], {
    x: 1.2, y: 4.55, w: 8.2, h: 0.9,
    fontFace: "Calibri", color: C.text, valign: "middle", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 11 — UC: What Makes It Different
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.tealDark } });
  slide.addText("UC Application: What Makes It Different", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 26, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const ucDiffs = [
    { title: "No Recommendation Letters", desc: "Most applicants don't submit letters. Your essays and activities speak for you.", color: C.teal },
    { title: "No Supplemental Essays", desc: "PIQs replace school-specific supplements. One set of essays for all 9 UC campuses.", color: C.tealDark },
    { title: "UC GPA Calculated Differently", desc: "Weighted GPA using 10th-11th grade courses only. Honors/AP courses get a bonus point (capped at 8 semesters).", color: C.teal },
    { title: "One Application, 9 Campuses", desc: "Apply to multiple UCs with a single application. Rank your campus/major preferences.", color: C.tealDark },
  ];
  for (let i = 0; i < 4; i++) {
    const y = 1.15 + i * 1.15;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.95,
      fill: { color: C.white }, shadow: cardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.08, h: 0.95,
      fill: { color: ucDiffs[i].color }
    });
    slide.addText(ucDiffs[i].title, {
      x: 0.8, y: y + 0.08, w: 8.5, h: 0.35,
      fontSize: 17, fontFace: "Calibri", bold: true,
      color: ucDiffs[i].color, margin: 0
    });
    slide.addText(ucDiffs[i].desc, {
      x: 0.8, y: y + 0.45, w: 8.5, h: 0.4,
      fontSize: 14, fontFace: "Calibri",
      color: C.text, margin: 0
    });
  }

  // ════════════════════════════════════════════════
  // SLIDE 12 — UC: 13 Review Criteria
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.tealDark } });
  slide.addText("UC Holistic Review: The 13 Criteria", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 26, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const criteria = [
    { num: "1", text: "GPA in all completed A-G courses", cat: "Education" },
    { num: "2", text: "GPA in A-G courses in last 2 years", cat: "Education" },
    { num: "3", text: "Number & rigor of courses beyond minimum", cat: "Education" },
    { num: "4", text: "Number of UC-approved honors/AP/IB/college courses", cat: "Education" },
    { num: "5", text: "Identification with underrepresented groups", cat: "Experience" },
    { num: "6", text: "Achievement in special projects, programs, activities", cat: "Exposure" },
    { num: "7", text: "Academic accomplishments in light of life experiences", cat: "Experience" },
    { num: "8", text: "Academic accomplishments in context of school", cat: "Education" },
    { num: "9", text: "Outstanding performance in academic subject areas", cat: "Education" },
    { num: "10", text: "Outstanding work in creative fields", cat: "Exposure" },
    { num: "11", text: "Demonstrated leadership ability", cat: "Experience" },
    { num: "12", text: "Community service and public engagement", cat: "Experience" },
    { num: "13", text: "Demonstrated interest and ability in a major", cat: "Exposure" },
  ];

  const catColors = { Education: C.teal, Experience: C.accent, Exposure: "7B1FA2" };

  // Table layout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 9.4, h: 3.9,
    fill: { color: C.white }, shadow: softShadow()
  });

  for (let i = 0; i < 13; i++) {
    const y = 1.08 + i * 0.295;
    const bg = i % 2 === 0 ? C.cream : C.white;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: 0.29,
      fill: { color: bg }
    });
    // Number
    slide.addText(criteria[i].num, {
      x: 0.4, y, w: 0.5, h: 0.29,
      fontSize: 10, fontFace: "Calibri", bold: true,
      color: C.teal, align: "center", valign: "middle", margin: 0
    });
    // Text
    slide.addText(criteria[i].text, {
      x: 0.9, y, w: 7.0, h: 0.29,
      fontSize: 10, fontFace: "Calibri",
      color: C.text, valign: "middle", margin: 0
    });
    // Category tag
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 8.1, y: y + 0.03, w: 1.4, h: 0.23,
      fill: { color: catColors[criteria[i].cat] }
    });
    slide.addText(criteria[i].cat, {
      x: 8.1, y: y + 0.03, w: 1.4, h: 0.23,
      fontSize: 9, fontFace: "Calibri", bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0
    });
  }

  // Note — positioned below table with clear gap
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.0, w: 9, h: 0.5,
    fill: { color: C.tealLight }
  });
  slide.addText("Notice how all 13 criteria map to the 3E framework: Education, Experience, Exposure", {
    x: 0.5, y: 5.0, w: 9, h: 0.5,
    fontSize: 13, fontFace: "Calibri", italic: true, bold: true,
    color: C.teal, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 13 — Comparison Chart
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addImage({ data: icons.compare, x: 0.4, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("Common App vs. UC: Side-by-Side", {
    x: 1.0, y: 0.1, w: 8.5, h: 0.7,
    fontSize: 26, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const compTable = [
    [
      { text: "Category", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 13, align: "center", valign: "middle" } },
      { text: "Common App", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 13, align: "center", valign: "middle" } },
      { text: "UC Application", options: { fill: { color: C.tealDark }, color: C.white, bold: true, fontSize: 13, align: "center", valign: "middle" } },
    ],
    [
      { text: "Schools", options: { bold: true, fontSize: 12, fill: { color: C.white } } },
      { text: "900+ nationwide", options: { fontSize: 12, fill: { color: C.white } } },
      { text: "9 UC campuses", options: { fontSize: 12, fill: { color: C.white } } },
    ],
    [
      { text: "Activities", options: { bold: true, fontSize: 12, fill: { color: C.tableAlt } } },
      { text: "10 slots, 150 chars", options: { fontSize: 12, fill: { color: C.tableAlt } } },
      { text: "20 slots + 5 awards", options: { fontSize: 12, fill: { color: C.tableAlt } } },
    ],
    [
      { text: "Essays", options: { bold: true, fontSize: 12, fill: { color: C.white } } },
      { text: "1 personal (650 words)\n+ supplementals", options: { fontSize: 12, fill: { color: C.white } } },
      { text: "4 PIQs (350 words each)", options: { fontSize: 12, fill: { color: C.white } } },
    ],
    [
      { text: "Recommendations", options: { bold: true, fontSize: 12, fill: { color: C.tableAlt } } },
      { text: "1-3 required (most schools)", options: { fontSize: 12, fill: { color: C.tableAlt } } },
      { text: "Not required", options: { fontSize: 12, fill: { color: C.tableAlt } } },
    ],
    [
      { text: "Test Scores", options: { bold: true, fontSize: 12, fill: { color: C.white } } },
      { text: "School-dependent", options: { fontSize: 12, fill: { color: C.white } } },
      { text: "Test-free (not considered)", options: { fontSize: 12, fill: { color: C.white } } },
    ],
    [
      { text: "GPA", options: { bold: true, fontSize: 12, fill: { color: C.tableAlt } } },
      { text: "School-reported", options: { fontSize: 12, fill: { color: C.tableAlt } } },
      { text: "UC-calculated (10th-11th)", options: { fontSize: 12, fill: { color: C.tableAlt } } },
    ],
    [
      { text: "Deadline", options: { bold: true, fontSize: 12, fill: { color: C.white } } },
      { text: "Varies (Nov 1 - Jan 15)", options: { fontSize: 12, fill: { color: C.white } } },
      { text: "November 30", options: { fontSize: 12, fill: { color: C.white } } },
    ],
  ];

  slide.addTable(compTable, {
    x: 0.5, y: 1.1, w: 9,
    colW: [2.2, 3.4, 3.4],
    border: { pt: 0.5, color: C.border },
    fontFace: "Calibri", color: C.text,
    rowH: [0.45, 0.42, 0.42, 0.52, 0.42, 0.42, 0.42, 0.42],
    valign: "middle",
  });

  // Discussion
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.8, w: 9, h: 0.6,
    fill: { color: C.accent }
  });
  slide.addImage({ data: icons.lightbulbWhite, x: 0.7, y: 4.88, w: 0.35, h: 0.35 });
  slide.addText("Discussion: Given these differences, what should you prepare differently for each?", {
    x: 1.2, y: 4.8, w: 8.1, h: 0.6,
    fontSize: 15, fontFace: "Calibri", bold: true, italic: true,
    color: C.white, valign: "middle", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 14 — Timeline
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addImage({ data: icons.calendarWhite, x: 0.4, y: 0.15, w: 0.5, h: 0.5 });
  slide.addText("The Timeline: When to Start What", {
    x: 1.0, y: 0.1, w: 8.5, h: 0.7,
    fontSize: 26, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Timeline
  const timeline = [
    { grade: "9th-10th", label: "Build Foundation", items: "Explore interests, join activities, keep grades strong, build reading habit", color: C.tealLight },
    { grade: "Summer\nbefore 11th", label: "Go Deeper", items: "Start leadership roles, enter competitions, begin research or passion projects", color: C.greenLight },
    { grade: "11th Grade", label: "Prepare", items: "Take SAT/ACT, build activity list, brainstorm essay topics, visit campuses", color: C.tealLight },
    { grade: "Summer\nbefore 12th", label: "Write & Refine", items: "Draft Common App essay + PIQs, write activity descriptions, finalize supplementals", color: C.greenLight },
    { grade: "12th Fall", label: "Submit!", items: "Polish essays, submit applications, request recommendations, send test scores", color: C.tealLight },
  ];

  // Connecting line
  slide.addShape(pres.shapes.LINE, {
    x: 1.3, y: 1.3, w: 0, h: 3.6,
    line: { color: C.teal, width: 3 }
  });

  for (let i = 0; i < 5; i++) {
    const y = 1.1 + i * 0.78;
    // Circle marker
    slide.addShape(pres.shapes.OVAL, {
      x: 1.05, y: y + 0.12, w: 0.5, h: 0.5,
      fill: { color: C.teal }
    });
    slide.addText(String(i + 1), {
      x: 1.05, y: y + 0.12, w: 0.5, h: 0.5,
      fontSize: 14, fontFace: "Calibri", bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0
    });
    // Content card
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.8, y, w: 7.8, h: 0.7,
      fill: { color: timeline[i].color }
    });
    slide.addText(timeline[i].grade, {
      x: 1.9, y, w: 1.4, h: 0.7,
      fontSize: 11, fontFace: "Calibri", bold: true,
      color: C.tealDark, valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.LINE, {
      x: 3.3, y: y + 0.1, w: 0, h: 0.5,
      line: { color: C.border, width: 1 }
    });
    slide.addText(timeline[i].label, {
      x: 3.4, y, w: 1.5, h: 0.35,
      fontSize: 13, fontFace: "Calibri", bold: true,
      color: C.teal, margin: 0
    });
    slide.addText(timeline[i].items, {
      x: 3.4, y: y + 0.32, w: 6.0, h: 0.35,
      fontSize: 11, fontFace: "Calibri",
      color: C.text, margin: 0
    });
  }

  // Bottom callout
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.0, w: 9, h: 0.5,
    fill: { color: C.accent }
  });
  slide.addText("You can't cram a 4-year story in senior fall. Start now.", {
    x: 0.5, y: 5.0, w: 9, h: 0.5,
    fontSize: 16, fontFace: "Calibri", bold: true,
    color: C.white, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 15 — 3E Self-Audit Worksheet
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.teal } });
  slide.addText("3E Self-Audit", {
    x: 0.5, y: 0.05, w: 4, h: 0.35,
    fontSize: 14, fontFace: "Calibri", charSpacing: 3,
    color: C.accent, margin: 0
  });
  slide.addText("Where Do You Stand?  (10-minute activity)", {
    x: 0.5, y: 0.35, w: 9, h: 0.5,
    fontSize: 22, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  // Audit table
  const auditTable = [
    [
      { text: "", options: { fill: { color: C.white }, border: { pt: 0.5, color: C.border } } },
      { text: "Education", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 14, align: "center" } },
      { text: "Experience", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 14, align: "center" } },
      { text: "Exposure", options: { fill: { color: C.teal }, color: C.white, bold: true, fontSize: 14, align: "center" } },
    ],
    [
      { text: "What I\nCurrently Have", options: { bold: true, fontSize: 12, fill: { color: C.tealLight }, color: C.tealDark } },
      { text: "", options: { fill: { color: C.white } } },
      { text: "", options: { fill: { color: C.white } } },
      { text: "", options: { fill: { color: C.white } } },
    ],
    [
      { text: "My Gaps\n(What's Missing?)", options: { bold: true, fontSize: 12, fill: { color: C.tealLight }, color: C.tealDark } },
      { text: "", options: { fill: { color: "FFF8E1" } } },
      { text: "", options: { fill: { color: "FFF8E1" } } },
      { text: "", options: { fill: { color: "FFF8E1" } } },
    ],
    [
      { text: "Action Items\n(This Summer / Next Year)", options: { bold: true, fontSize: 12, fill: { color: C.tealLight }, color: C.tealDark } },
      { text: "", options: { fill: { color: C.greenLight } } },
      { text: "", options: { fill: { color: C.greenLight } } },
      { text: "", options: { fill: { color: C.greenLight } } },
    ],
  ];

  slide.addTable(auditTable, {
    x: 0.4, y: 1.1, w: 9.2,
    colW: [2.0, 2.4, 2.4, 2.4],
    rowH: [0.5, 1.1, 1.1, 1.1],
    border: { pt: 1, color: C.border },
    fontFace: "Calibri", color: C.text, valign: "middle",
  });

  slide.addText("Take notes! You'll need this for your reflection.", {
    x: 0.5, y: 5.1, w: 9, h: 0.35,
    fontSize: 13, fontFace: "Calibri", italic: true,
    color: C.textMid, align: "center", margin: 0
  });

  // ════════════════════════════════════════════════
  // SLIDE 16 — Reflection: Where Am I Now?
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.tealDark };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

  slide.addText("REFLECTION", {
    x: 0.5, y: 0.3, w: 9, h: 0.35,
    fontSize: 14, fontFace: "Calibri", charSpacing: 4,
    color: C.accent, margin: 0
  });
  slide.addText("Where Am I Now?", {
    x: 0.5, y: 0.65, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const reflectQ1 = [
    { q: "Which 3E category is your strongest?\nWhich has the biggest gap?", num: "01" },
    { q: "List 3 activities you could put on\nan application today.", num: "02" },
    { q: "If you had to write your Common App essay\nright now, which prompt would you pick? Why?", num: "03" },
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.5 + i * 1.3;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 1.1,
      fill: { color: C.teal }
    });
    slide.addText(reflectQ1[i].num, {
      x: 0.7, y: y + 0.1, w: 0.8, h: 0.8,
      fontSize: 32, fontFace: "Georgia", bold: true,
      color: C.accent, valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.LINE, {
      x: 1.6, y: y + 0.15, w: 0, h: 0.8,
      line: { color: C.border, width: 1 }
    });
    slide.addText(reflectQ1[i].q, {
      x: 1.8, y, w: 7.5, h: 1.1,
      fontSize: 16, fontFace: "Calibri",
      color: C.white, valign: "middle", margin: 0
    });
  }

  // ════════════════════════════════════════════════
  // SLIDE 17 — Reflection: What's My Plan?
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.tealDark };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

  slide.addText("REFLECTION", {
    x: 0.5, y: 0.3, w: 9, h: 0.35,
    fontSize: 14, fontFace: "Calibri", charSpacing: 4,
    color: C.accent, margin: 0
  });
  slide.addText("What's My Plan?", {
    x: 0.5, y: 0.65, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const reflectQ2 = [
    { q: "Name one activity you can start or\ndeepen this summer.", num: "01" },
    { q: "What can you do this month to strengthen\nyour weakest 3E area?", num: "02" },
    { q: "Who could write you a strong recommendation\nletter? What do they know about you?", num: "03" },
    { q: "Set a goal:\nBy _______, I will have _____________.", num: "04" },
  ];
  for (let i = 0; i < 4; i++) {
    const y = 1.4 + i * 1.0;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.85,
      fill: { color: C.teal }
    });
    slide.addText(reflectQ2[i].num, {
      x: 0.7, y: y + 0.05, w: 0.8, h: 0.7,
      fontSize: 28, fontFace: "Georgia", bold: true,
      color: C.accent, valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.LINE, {
      x: 1.6, y: y + 0.1, w: 0, h: 0.65,
      line: { color: C.border, width: 1 }
    });
    slide.addText(reflectQ2[i].q, {
      x: 1.8, y, w: 7.5, h: 0.85,
      fontSize: 15, fontFace: "Calibri",
      color: C.white, valign: "middle", margin: 0
    });
  }

  // ════════════════════════════════════════════════
  // SLIDE 18 — Key Takeaways & Next Steps
  // ════════════════════════════════════════════════
  slide = pres.addSlide();
  slide.background = { color: C.tealDark };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

  slide.addText("Key Takeaways", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 32, fontFace: "Georgia", bold: true,
    color: C.white, margin: 0
  });

  const takeaways = [
    { icon: icons.checkWhite, text: "Know the platforms — Common App and UC have different requirements, timelines, and strategies." },
    { icon: icons.pencil, text: "Essays and activities need years of preparation, not weeks. Start writing early and revise often." },
    { icon: icons.rocketWhite, text: "Use the 3E framework (Education, Experience, Exposure) to audit yourself regularly." },
    { icon: icons.starWhite, text: "Start now — the best application tells a story built over time, not assembled last minute." },
  ];
  for (let i = 0; i < 4; i++) {
    const y = 1.2 + i * 0.95;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.75,
      fill: { color: C.teal }
    });
    slide.addImage({ data: takeaways[i].icon, x: 0.7, y: y + 0.12, w: 0.5, h: 0.5 });
    slide.addText(takeaways[i].text, {
      x: 1.4, y, w: 7.9, h: 0.75,
      fontSize: 15, fontFace: "Calibri",
      color: C.white, valign: "middle", margin: 0
    });
  }

  // Footer
  slide.addText("GR EDU  |  谷雨中文", {
    x: 0.5, y: 5.0, w: 4, h: 0.4,
    fontSize: 14, fontFace: "Calibri",
    color: C.border, margin: 0
  });
  slide.addText("Questions?", {
    x: 5.5, y: 5.0, w: 4, h: 0.4,
    fontSize: 18, fontFace: "Georgia", bold: true, italic: true,
    color: C.accent, align: "right", margin: 0
  });

  // ── SAVE ──
  await pres.writeFile({ fileName: "English/college_app_platforms.pptx" });
  console.log("Created: English/college_app_platforms.pptx (18 slides)");
}

main().catch(err => { console.error(err); process.exit(1); });

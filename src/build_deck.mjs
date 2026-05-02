/** @jsxRuntime automatic */
/** @jsxImportSource @oai/artifact-tool/presentation-jsx */

const {
  Presentation,
  PresentationFile,
  drawSlideToCtx,
  row,
  column,
  grid,
  panel,
  text,
  rule,
  fill,
  hug,
  fixed,
  wrap,
  grow,
  fr,
  auto,
} = await import("@oai/artifact-tool");
const { createRequire } = await import("node:module");
const fs = await import("node:fs");

const require = createRequire(import.meta.url);
const { Canvas, loadImage } = require("../node_modules/@oai/artifact-tool/node_modules/skia-canvas");

const OUT = "output/output.pptx";

const C = {
  paper: "#F7F3EC",
  ink: "#2A241E",
  ink2: "#5D5247",
  ink3: "#887C70",
  ember: "#B25A32",
  ember2: "#D18455",
  rule: "#C8BAA8",
  ivory: "#FFFDF8",
  dark: "#191612",
  green: "#607466",
  blue: "#536B80",
};

const presentation = Presentation.create({
  slideSize: { width: 1920, height: 1080 },
});

const base = {
  title: { fontFace: "Georgia", fontSize: 68, color: C.ink, bold: false },
  h1: { fontFace: "Georgia", fontSize: 56, color: C.ink, bold: false },
  h2: { fontFace: "Georgia", fontSize: 38, color: C.ink, bold: false },
  body: { fontFace: "Aptos", fontSize: 27, color: C.ink2 },
  bodyDark: { fontFace: "Aptos", fontSize: 27, color: C.paper },
  mono: { fontFace: "Aptos Mono", fontSize: 16, color: C.ember, bold: true },
  label: { fontFace: "Aptos", fontSize: 18, color: C.ember, bold: true },
  small: { fontFace: "Aptos", fontSize: 16, color: C.ink3 },
};

function compose(slide, node) {
  slide.compose(node, {
    frame: { left: 0, top: 0, width: 1920, height: 1080 },
    baseUnit: 8,
  });
}

function addSlide(children, opts = {}) {
  const slide = presentation.slides.add();
  const bg = opts.bg || C.paper;
  compose(
    slide,
    panel(
      {
        name: "slide-bg",
        width: fill,
        height: fill,
        fill: bg,
        padding: 0,
      },
      column(
        {
          name: "slide-shell",
          width: fill,
          height: fill,
          padding: { x: 86, y: 60 },
          gap: 26,
        },
        children,
      ),
    ),
  );
  return slide;
}

function meta(left, num, dark = false) {
  return row(
    { name: "meta", width: fill, height: hug, align: "center", justify: "between" },
    [
      row({ name: "meta-left", width: hug, height: hug, gap: 12, align: "center" }, [
        panel({ name: "dot", width: fixed(12), height: fixed(12), fill: dark ? C.ember2 : C.ember, borderRadius: "rounded-full" }),
        text(left, { name: "meta-label", width: hug, height: hug, style: { ...base.mono, color: dark ? C.ember2 : C.ember } }),
      ]),
      text(String(num).padStart(2, "0"), {
        name: "slide-num",
        width: hug,
        height: hug,
        style: { ...base.mono, color: dark ? "#BFB4A5" : C.ink3 },
      }),
    ],
  );
}

function footer(section, num, dark = false) {
  return column({ name: "footer", width: fill, height: hug, gap: 12 }, [
    rule({ name: "footer-rule", width: fill, stroke: dark ? "#4A4037" : C.rule, weight: 1 }),
    row({ name: "footer-row", width: fill, height: hug, justify: "between" }, [
      text(section, { width: hug, height: hug, style: { ...base.small, color: dark ? "#BFB4A5" : C.ink3 } }),
      text(`${String(num).padStart(2, "0")} / 22`, { width: hug, height: hug, style: { ...base.small, color: dark ? "#BFB4A5" : C.ink3 } }),
    ]),
  ]);
}

function spacer(weight = 1) {
  return panel({ name: `spacer-${weight}`, width: fill, height: fixed(Math.round(54 * weight)), materialize: false });
}

function bulletList(items, style = base.body) {
  return column(
    { name: "bullets", width: fill, height: hug, gap: 12 },
    items.map((item, i) =>
      row({ name: `bullet-${i + 1}`, width: fill, height: hug, gap: 14, align: "start" }, [
        text("-", { width: fixed(24), height: hug, style: { ...style, color: C.ember, bold: true } }),
        text(item, { width: fill, height: hug, style }),
      ]),
    ),
  );
}

function titleBlock(title, subtitle, maxWidth = 1420) {
  return column({ name: "title-block", width: wrap(maxWidth), height: hug, gap: 18 }, [
    text(title, { name: "slide-title", width: fill, height: hug, style: base.h1 }),
    subtitle
      ? text(subtitle, { name: "subtitle", width: wrap(maxWidth - 80), height: hug, style: { ...base.body, fontSize: 28, color: C.ink2 } })
      : text("", { width: fixed(1), height: fixed(1), style: { fontSize: 1, color: C.paper } }),
  ]);
}

function sectionSlide(num, part, title, kicker, footerText) {
  addSlide(
    [
      meta(part, num),
      spacer(1),
      column({ name: "section-main", width: wrap(1500), height: hug, gap: 24 }, [
        rule({ width: fixed(110), stroke: C.ember, weight: 8 }),
        text(part.toUpperCase(), { width: hug, height: hug, style: { ...base.mono, fontSize: 19, letterSpacing: 1.3 } }),
        text(title, { width: fill, height: hug, style: { ...base.title, fontSize: 76 } }),
        text(kicker, { width: wrap(1320), height: hug, style: { ...base.body, fontSize: 31, color: C.ink2 } }),
      ]),
      spacer(1),
      footer(footerText, num),
    ],
    { bg: C.paper },
  );
}

// 01
addSlide([
  spacer(0.7),
  column({ name: "cover", width: fill, height: hug, gap: 32 }, [
    rule({ width: fixed(150), stroke: C.ember, weight: 8 }),
    text("Friction.", { width: fill, height: hug, style: { ...base.title, fontSize: 132 } }),
    text("THE PEDAGOGICAL IMPERATIVE OF FRICTION", {
      width: fill,
      height: hug,
      style: { ...base.mono, fontSize: 24, letterSpacing: 1.8 },
    }),
    text("Why productive struggle matters more than ever in K-12 in the age of generative AI.", {
      width: wrap(1280),
      height: hug,
      style: { ...base.body, fontSize: 34, color: C.ink2 },
    }),
  ]),
  spacer(1),
  row({ name: "cover-bottom", width: fill, height: hug, justify: "between", align: "end" }, [
    column({ width: hug, height: hug, gap: 6 }, [
      text("Micah J. Miner, CETL, Ed.S.", { width: hug, height: hug, style: { ...base.body, color: C.ink, bold: true } }),
      text("Director of Innovation & Technology  -  Doctoral Candidate", { width: hug, height: hug, style: { ...base.small, fontSize: 19 } }),
    ]),
    text("ISTE+ASCD Webinar  -  2026", { width: hug, height: hug, style: base.small }),
  ]),
  footer("A media-ecological case for K-12 leaders", 1),
]);

// 02
addSlide([
  meta("Opening  -  The framing problem", 2),
  titleBlock("The wrong questions are louder than the right one.", "The work is not only to manage tools. It is to understand the learning environment children now inhabit."),
  grid(
    { name: "wrong-right", width: fill, height: grow(1), columns: [fr(0.9), fr(1.1)], columnGap: 56, align: "center" },
    [
      column({ width: fill, height: hug, gap: 24 }, [
        text("Common district questions", { width: fill, height: hug, style: { ...base.label, fontSize: 22, color: C.ink3 } }),
        bulletList(["Should we ban AI?", "How do we catch misuse?", "Which vendor should we buy?"], { ...base.body, fontSize: 32, color: C.ink2 }),
      ]),
      panel(
        { width: fill, height: hug, padding: { x: 42, y: 36 }, fill: C.dark, borderRadius: 8 },
        column({ width: fill, height: hug, gap: 14 }, [
          text("The better question", { width: fill, height: hug, style: { ...base.label, fontSize: 22, color: C.ember2 } }),
          text("What happens to learning when the symbolic environment children think inside is increasingly written by machines?", {
            width: fill,
            height: hug,
            style: { ...base.h1, fontSize: 46, color: C.paper },
          }),
        ]),
      ),
    ],
  ),
  footer("Section I  -  Reframing", 2),
]);

// 03
addSlide([
  meta("Two ways of framing the problem", 3),
  titleBlock("A policy question, or a question of consciousness.", "Both matter. Only one explains why generative AI changes the developmental conditions of school."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1)], columnGap: 46, align: "stretch" }, [
    panel({ padding: { x: 36, y: 34 }, fill: C.ivory, borderRadius: 8 }, column({ width: fill, height: fill, gap: 18 }, [
      text("Framing A  -  Tool adoption", { width: fill, height: hug, style: { ...base.h2, color: C.ink } }),
      text("AI is a productivity tool. Manage it through integrity policies, data privacy, vendor review, and acceptable use.", { width: fill, height: hug, style: base.body }),
      spacer(1),
      text("Necessary, but incomplete.", { width: fill, height: hug, style: { ...base.label, color: C.ember } }),
    ])),
    panel({ padding: { x: 36, y: 34 }, fill: C.dark, borderRadius: 8 }, column({ width: fill, height: fill, gap: 18 }, [
      text("Framing B  -  Media ecology", { width: fill, height: hug, style: { ...base.h2, color: C.paper } }),
      text("AI is an environment. Each new medium restructures how children remember, attend, reason, compose, and know.", { width: fill, height: hug, style: { ...base.bodyDark, color: "#E7DED2" } }),
      spacer(1),
      text("This is the frame K-12 needs next.", { width: fill, height: hug, style: { ...base.label, color: C.ember2 } }),
    ])),
  ]),
  footer("Section I  -  Reframing", 3),
]);

// 04
addSlide([
  meta("What we'll do in the next 30 minutes", 4),
  titleBlock("Roadmap.", "Four moves from theory to Monday morning practice."),
  column({ width: fill, height: grow(1), gap: 20, justify: "center" }, [
    ["01", "Stages of the symbolic environment", "From oral cultures to algorithmic generation, and where K-12 actually sits."],
    ["02", "What this does to learning", "Productive struggle, unproductive success, and the bypass of cognition."],
    ["03", "Pedagogical friction", "Four dimensions of the response framework: head, room, world, system."],
    ["04", "What changes Monday morning", "Curriculum, advocacy, and policy moves for K-12 leaders."],
  ].map(([n, h, b]) => row({ width: fill, height: hug, gap: 28, align: "start" }, [
    text(n, { width: fixed(70), height: hug, style: { ...base.mono, fontSize: 28 } }),
    column({ width: fill, height: hug, gap: 4 }, [
      text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 34 } }),
      text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 24 } }),
    ]),
  ])),
  ),
  footer("Section I  -  Reframing", 4),
]);

sectionSlide(5, "Part One  -  Media Ecology", "Stages of the symbolic environment.", "Each new medium reorganizes the cognition that lives inside it. K-12 is now negotiating the first transition in which content itself is generated by machines.", "Walter Ong  -  post-Ong media ecology");

// 06
addSlide([
  meta("The developmental account, extended", 6),
  titleBlock("Five stages, not three.", "Ong gave us oral, literate, and broadcast cultures. The current moment requires two further stages."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1), fr(1), fr(1), fr(1)], columnGap: 0, align: "stretch" }, [
    ["01", "Primary orality", "Memory-based, communal, formulaic.", "PRE-LITERATE"],
    ["02", "Literacy", "Writing externalizes memory and supports abstraction.", "PRINT ERA"],
    ["03", "Secondary orality", "Broadcast retrieves orality inside literate infrastructure.", "RADIO  -  TV"],
    ["04", "Algorithmic secondary orality", "Humans author; algorithms decide what reaches whom.", "SOCIAL  -  2005-"],
    ["05", "Tertiary algorithmicity", "Machines both curate and generate symbolic content.", "GENAI  -  2022-"],
  ].map(([n, h, b, era], i) =>
    panel({ padding: { x: 20, y: 24 }, fill: i === 4 ? C.dark : "transparent" }, column({ width: fill, height: fill, gap: 14 }, [
      text(`STAGE ${n}`, { width: fill, height: hug, style: { ...base.mono, color: i === 4 ? C.ember2 : C.ember } }),
      text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 31, color: i === 4 ? C.paper : C.ink } }),
      text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 22, color: i === 4 ? "#E5D9CA" : C.ink2 } }),
      spacer(1),
      text(era, { width: fill, height: hug, style: { ...base.small, color: i === 4 ? "#BFB4A5" : C.ink3 } }),
  ])),
  )),
  footer("Section II  -  Theoretical framework", 6),
]);

// 07
addSlide([
  meta("Why we need a new stage at all", 7),
  titleBlock("Three assumptions no longer hold.", "The media environment now changes origin, distribution, and the learner's relation to the world."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1), fr(1)], columnGap: 34, align: "stretch" }, [
    ["01", "Humans originate the content.", "Neural networks now produce essays, summaries, analyses, and images without human authorship in the traditional sense."],
    ["02", "Distribution is transparent.", "Two students opening the same app may see different content, curated by opaque systems."],
    ["03", "Media remains external.", "Personalized environments can reflect a child's prior behavior back, reducing the disequilibrium learning requires."],
  ].map(([n, h, b]) => panel({ padding: { x: 34, y: 34 }, fill: C.ivory, borderRadius: 8 }, column({ width: fill, height: fill, gap: 18 }, [
    text(n, { width: fill, height: hug, style: { ...base.mono, fontSize: 22 } }),
    text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 34 } }),
    text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 24 } }),
  ])))),
  footer("Section II  -  Theoretical framework", 7),
]);

// 08
addSlide([
  meta("The proposed stage, in one sentence", 8),
  spacer(1),
  column({ width: fixed(1500), height: hug, gap: 26 }, [
    rule({ width: fixed(120), stroke: C.ember, weight: 8 }),
    text("Tertiary algorithmicity (n.)", { width: fill, height: hug, style: { ...base.mono, fontSize: 24 } }),
    text("The condition in which algorithmic systems both curate and generate the symbolic content children learn inside.", {
      width: fixed(1450),
      height: hug,
      style: { ...base.title, fontSize: 72 },
    }),
    text("This makes human authorship optional at scale.", { width: fill, height: hug, style: { ...base.body, fontSize: 34, color: C.ember } }),
  ]),
  spacer(1),
  footer("Section III  -  Proposed extensions", 8),
]);

// 09
addSlide([
  meta("What's new about this stage, specifically", 9),
  titleBlock("Three characteristics of the new stage.", "Each one names a pressure point for teaching, learning, and governance."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1), fr(1)], columnGap: 36 }, [
    ["Noetic displacement", "Meaning-making can be performed by a system that has never had an experience."],
    ["Rhetorical saturation", "Fluent prose floods every channel, weakening the signal of competent thought."],
    ["Existential abstraction", "The link between a claim and the person making it becomes less stable."],
  ].map(([h, b]) => column({ width: fill, height: fill, gap: 18, justify: "center" }, [
    rule({ width: fixed(90), stroke: C.ember, weight: 5 }),
    text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 38 } }),
    text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 27 } }),
  ]))),
  footer("Section III  -  Proposed extensions", 9),
]);

sectionSlide(10, "Part Two  -  Learning Science", "What this does to a child who is learning.", "Durable learning requires productive struggle. Tertiary algorithmicity makes that struggle optional.", "Dewey  -  Piaget  -  Vygotsky  -  Kapur  -  Bjork  -  Sweller");

// 11
addSlide([
  meta("Kapur's four quadrants", 11),
  titleBlock("Performance is not the same as learning.", "Kapur's 2x2 names the precise risk of generative AI in K-12."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1)], rows: [fr(1), fr(1)], columnGap: 20, rowGap: 20 }, [
    ["Quadrant of concern", "Unproductive success", "Correct output. No schema construction. The artifact looks like learning; the learning never occurred.", C.dark, C.paper],
    ["Goal state", "Productive success", "Correct output and genuine understanding. The student wrestled with the material.", C.ivory, C.ink],
    ["Avoid", "Unproductive failure", "Wrong output, no growth. The task exceeded reach without scaffolding.", C.ivory, C.ink],
    ["Often the path to mastery", "Productive failure", "Wrong output, real learning. Struggle activates prior knowledge and prepares future learning.", "#EFE7DC", C.ink],
  ].map(([label, h, b, bg, fg]) => panel({ padding: { x: 28, y: 24 }, fill: bg, borderRadius: 8 }, column({ width: fill, height: fill, gap: 12 }, [
    text(label, { width: fill, height: hug, style: { ...base.label, color: bg === C.dark ? C.ember2 : C.ember } }),
    text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 34, color: fg } }),
    text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 22, color: bg === C.dark ? "#E8DED2" : C.ink2 } }),
  ])))),
  footer("Section IV  -  Educational stakes", 11),
]);

// 12
addSlide([
  meta("The central educational threat", 12),
  spacer(1),
  column({ width: fixed(1500), height: hug, gap: 24 }, [
    rule({ width: fixed(120), stroke: C.ember, weight: 8 }),
    text("Unproductive success", { width: fill, height: hug, style: { ...base.mono, fontSize: 24 } }),
    text("The artifact of learning remains visible. The process of learning becomes optional.", {
      width: fixed(1450),
      height: hug,
      style: { ...base.title, fontSize: 76 },
    }),
    text("After Kapur (2016), applied to K-12 under tertiary algorithmicity.", { width: fill, height: hug, style: { ...base.small, fontSize: 20 } }),
  ]),
  spacer(1),
  footer("Section IV  -  Educational stakes", 12),
]);

// 13
addSlide([
  meta("A common mistake in K-12 framing", 13),
  titleBlock("This is not, primarily, an academic integrity problem.", "Integrity matters, but it is not the deepest educational diagnosis."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1)], columnGap: 42 }, [
    panel({ padding: { x: 34, y: 34 }, fill: C.ivory, borderRadius: 8 }, column({ width: fill, height: fill, gap: 18 }, [
      text("The integrity frame", { width: fill, height: hug, style: { ...base.h2, fontSize: 38 } }),
      text("Locates the problem in students who cheat. Solves it with detection tools, honor codes, and locked browsers.", { width: fill, height: hug, style: base.body }),
      spacer(1),
      text("Necessary, but aimed at too narrow a target.", { width: fill, height: hug, style: { ...base.label } }),
    ])),
    panel({ padding: { x: 34, y: 34 }, fill: C.dark, borderRadius: 8 }, column({ width: fill, height: fill, gap: 18 }, [
      text("The developmental frame", { width: fill, height: hug, style: { ...base.h2, fontSize: 38, color: C.paper } }),
      text("Locates the problem in the media environment itself, whose default tendency is to bypass cognitive labor.", { width: fill, height: hug, style: { ...base.bodyDark } }),
      spacer(1),
      text("Students are not lazy. The path of least resistance changed.", { width: fill, height: hug, style: { ...base.label, color: C.ember2 } }),
    ])),
  ]),
  footer("Section IV  -  Educational stakes", 13),
]);

// 14
addSlide([
  meta("What's at risk inside the K-12 classroom", 14),
  titleBlock("Three losses, K-12 specific.", "The stakes are developmental before they are disciplinary."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1), fr(1)], columnGap: 36 }, [
    ["Attention & the sensorium", "Feeds reward rapid consumption over deliberation. Generated content adds fluency and volume."],
    ["Knowledge & interpretation", "Students must evaluate claims whose origin and authority are uncertain."],
    ["Agency & authorship", "Creative and epistemic agency become distributed between child and system."],
  ].map(([h, b]) => column({ width: fill, height: fill, gap: 18, justify: "center" }, [
    text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 38 } }),
    rule({ width: fixed(110), stroke: C.ember, weight: 5 }),
    text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 27 } }),
  ]))),
  footer("Section IV  -  Educational stakes", 14),
]);

sectionSlide(15, "Part Three  -  A Response Framework", "Pedagogical friction.", "Not a call to make school harder. A call to protect the specific difficulties that learning depends on when the medium defaults to removing them.", "Section V  -  Pedagogical friction");

// 16
addSlide([
  meta("Definition", 16),
  spacer(1),
  column({ width: fixed(1520), height: hug, gap: 24 }, [
    rule({ width: fixed(120), stroke: C.ember, weight: 8 }),
    text("Pedagogical friction (n.)", { width: fill, height: hug, style: { ...base.mono, fontSize: 24 } }),
    text("The intentional preservation of the cognitive, dialogic, and existential resistance that durable learning requires.", {
      width: fixed(1480),
      height: hug,
      style: { ...base.title, fontSize: 70 },
    }),
    text("In physics, friction is energy lost. In pedagogy, it is the energy through which development occurs.", {
      width: fill,
      height: hug,
      style: { ...base.body, fontSize: 31, color: C.ember },
    }),
  ]),
  spacer(1),
  footer("Section V  -  Pedagogical friction", 16),
]);

// 17
addSlide([
  meta("The framework, four dimensions", 17),
  titleBlock("Four dimensions of friction.", "Three phenomenological, one structural. Together they specify what K-12 has to protect."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1), fr(1), fr(1)], columnGap: 18 }, [
    ["The Head", "Noetic friction", "Comprehension, synthesis, and revising one's own thinking.", "Bypass"],
    ["The Room", "Rhetorical friction", "Defending claims to unpredictable human interlocutors.", "Simulation"],
    ["The World", "Existential friction", "Being accountable for one's thinking in front of others.", "Abstraction"],
    ["The System", "Infrastructural friction", "Policy, assessment design, professional development, and institutional values.", "Silence"],
  ].map(([where, h, b, threat]) => panel({ padding: { x: 22, y: 24 }, fill: C.ivory, borderRadius: 8 }, column({ width: fill, height: fill, gap: 12 }, [
    text(where, { width: fill, height: hug, style: { ...base.mono, fontSize: 17 } }),
    text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 31 } }),
    text(b, { width: fill, height: hug, style: { ...base.body, fontSize: 21 } }),
    spacer(1),
    rule({ width: fill, stroke: C.rule, weight: 1 }),
    text(`Threat  -  ${threat}`, { width: fill, height: hug, style: { ...base.label, fontSize: 18 } }),
  ])))),
  footer("Section V  -  Pedagogical friction", 17),
]);

// 18
addSlide([
  meta("The equity guardrail on the framework", 18),
  titleBlock("Not all friction is productive.", "\"Preserve struggle\" without this distinction can reproduce inequity under the language of rigor."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1)], columnGap: 46 }, [
    panel({ padding: { x: 38, y: 34 }, fill: C.ivory, borderRadius: 8 }, column({ width: fill, height: fill, gap: 20 }, [
      text("Preserve productive friction", { width: fill, height: hug, style: { ...base.h2, fontSize: 40, color: C.green } }),
      bulletList(["Wrestling with a difficult text", "Defending an argument under critique", "Composing original work from one's own understanding", "Revising in response to feedback"], { ...base.body, fontSize: 25 }),
    ])),
    panel({ padding: { x: 38, y: 34 }, fill: C.dark, borderRadius: 8 }, column({ width: fill, height: fill, gap: 20 }, [
      text("Reduce exclusionary friction", { width: fill, height: hug, style: { ...base.h2, fontSize: 40, color: C.ember2 } }),
      bulletList(["Language barriers unrelated to the objective", "Inaccessible formats", "Procedural demands that do not measure learning", "Norms that treat one cognitive profile as default"], { ...base.bodyDark, fontSize: 25 }),
    ])),
  ]),
  footer("Section V  -  Pedagogical friction", 18),
]);

// 19
addSlide([
  meta("A K-12 case the framework has to handle", 19),
  titleBlock("The English Learner paradox.", "A multilingual student uses AI translation to convert a draft from their first language into English. What just happened?"),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1)], columnGap: 44 }, [
    column({ width: fill, height: fill, gap: 18 }, [
      text("Access gained", { width: fill, height: hug, style: { ...base.h2, fontSize: 42, color: C.green } }),
      text("An exclusionary barrier dropped. The student can demonstrate content knowledge without being filtered through a language barrier unrelated to the learning objective.", { width: fill, height: hug, style: base.body }),
    ]),
    column({ width: fill, height: fill, gap: 18 }, [
      text("Voice at risk", { width: fill, height: hug, style: { ...base.h2, fontSize: 42, color: C.ember } }),
      text("Translation can produce grammatically correct, stylistically generic prose, flattening phrasing that made the writing recognizably the student's own.", { width: fill, height: hug, style: base.body }),
    ]),
  ]),
  text("No mechanical rule resolves this. The same tool can reduce both kinds of friction at once. The framework asks educators to make situated professional judgments.", {
    width: fill,
    height: hug,
    style: { ...base.body, fontSize: 25, color: C.ink },
  }),
  footer("Section V  -  Pedagogical friction", 19),
]);

sectionSlide(20, "Part Four  -  Implications", "What changes Monday morning.", "Curriculum, advocacy, and policy moves for K-12 leaders who are tired of choosing between ban and embrace.", "Section V.D  -  Implications for K-12");

// 21
addSlide([
  meta("Three columns of moves", 21),
  titleBlock("Curriculum, advocacy, policy.", "Friction must become structural, not personal."),
  grid({ width: fill, height: grow(1), columns: [fr(1), fr(1), fr(1)], columnGap: 28 }, [
    ["Curriculum", "Make thinking visible.", ["Assess process, not only product", "Require noetic struggle before AI assistance", "Bring back oral defense, critique, in-class drafting", "Ask students to explain how work was made"]],
    ["Advocacy", "Name friction as protected.", ["Push back on efficiency as the only frame", "Speak from technology leadership, not against it", "Equip teachers to explain why struggle matters", "Bring families into the conversation early"]],
    ["Policy", "Close the learning-science gap.", ["Reference cognition, not only risk", "Distinguish productive from exclusionary use cases", "Build assessment that values process", "Fund professional judgment, not just tool training"]],
  ].map(([area, h, items]) => panel({ padding: { x: 26, y: 28 }, fill: area === "Advocacy" ? C.dark : C.ivory, borderRadius: 8 }, column({ width: fill, height: fill, gap: 16 }, [
    text(area, { width: fill, height: hug, style: { ...base.mono, fontSize: 18, color: area === "Advocacy" ? C.ember2 : C.ember } }),
    text(h, { width: fill, height: hug, style: { ...base.h2, fontSize: 32, color: area === "Advocacy" ? C.paper : C.ink } }),
    bulletList(items, { ...(area === "Advocacy" ? base.bodyDark : base.body), fontSize: 20 }),
  ])))),
  footer("Section V.D  -  Implications for K-12", 21),
]);

// 22
addSlide([
  meta("Closing", 22),
  spacer(1),
  column({ width: fixed(1500), height: hug, gap: 28 }, [
    rule({ width: fixed(120), stroke: C.ember, weight: 8 }),
    text("The question to keep asking", { width: fill, height: hug, style: { ...base.mono, fontSize: 24 } }),
    text("For whom does this difficulty build capacity, and for whom does it restrict access?", {
      width: fixed(1450),
      height: hug,
      style: { ...base.title, fontSize: 76 },
    }),
    text("Thank you. Questions welcome.", { width: fill, height: hug, style: { ...base.body, fontSize: 34, color: C.ember } }),
  ]),
  spacer(1),
  row({ width: fill, height: hug, justify: "between", align: "end" }, [
    column({ width: hug, height: hug, gap: 6 }, [
      text("Micah J. Miner, CETL, Ed.S.", { width: hug, height: hug, style: { ...base.body, color: C.ink, bold: true } }),
      text("Director of Innovation & Technology  -  Doctoral Candidate, National Louis University", { width: hug, height: hug, style: { ...base.small, fontSize: 19 } }),
    ]),
    text("22 / 22", { width: hug, height: hug, style: base.small }),
  ]),
]);

const pptxBlob = await PresentationFile.exportPptx(presentation);
await pptxBlob.save(OUT);
console.log(`Saved ${OUT}`);

fs.mkdirSync("scratch/previews", { recursive: true });
const slides = presentation.slides.items;
for (let i = 0; i < slides.length; i += 1) {
  const canvas = new Canvas(1920, 1080);
  const ctx = canvas.getContext("2d");
  await drawSlideToCtx(slides[i], presentation, ctx);
  await canvas.toFile(`scratch/previews/slide-${String(i + 1).padStart(2, "0")}.png`);
}

const thumbW = 384;
const thumbH = 216;
const cols = 4;
const rows = Math.ceil(slides.length / cols);
const montage = new Canvas(cols * thumbW, rows * thumbH);
const mctx = montage.getContext("2d");
mctx.fillStyle = "#F7F3EC";
mctx.fillRect(0, 0, montage.width, montage.height);
for (let i = 0; i < slides.length; i += 1) {
  const img = await loadImage(`scratch/previews/slide-${String(i + 1).padStart(2, "0")}.png`);
  const x = (i % cols) * thumbW;
  const y = Math.floor(i / cols) * thumbH;
  mctx.drawImage(img, x, y, thumbW, thumbH);
}
await montage.toFile("scratch/previews/contact-sheet.png");
console.log("Saved scratch/previews/*.png");




/************************************************************
 * PRABHAS PPT MAKER - Professional Template Engine
 * Clean, modern slide designs with consistent styling
 ************************************************************/

// ========== COLOR PALETTES ==========
// Professional color combinations
const COLOR_PALETTES = [
  { primary: "1E40AF", secondary: "3B82F6", accent: "93C5FD", light: "DBEAFE", dark: "1E3A8A" }, // Blue
  { primary: "047857", secondary: "10B981", accent: "6EE7B7", light: "D1FAE5", dark: "064E3B" }, // Emerald
  { primary: "7C3AED", secondary: "8B5CF6", accent: "C4B5FD", light: "EDE9FE", dark: "5B21B6" }, // Purple
  { primary: "DC2626", secondary: "EF4444", accent: "FCA5A5", light: "FEE2E2", dark: "991B1B" }, // Red
  { primary: "D97706", secondary: "F59E0B", accent: "FCD34D", light: "FEF3C7", dark: "92400E" }, // Amber
  { primary: "0891B2", secondary: "06B6D4", accent: "67E8F9", light: "CFFAFE", dark: "155E75" }, // Cyan
  { primary: "BE185D", secondary: "EC4899", accent: "F9A8D4", light: "FCE7F3", dark: "831843" }, // Pink
  { primary: "4F46E5", secondary: "6366F1", accent: "A5B4FC", light: "E0E7FF", dark: "3730A3" }, // Indigo
];

let paletteIndex = 0;

function getNextPalette() {
  const palette = COLOR_PALETTES[paletteIndex % COLOR_PALETTES.length];
  paletteIndex++;
  return palette;
}

function resetColors() {
  paletteIndex = Math.floor(Math.random() * COLOR_PALETTES.length);
}

// Random color generator for dynamic template colors
// Rainbow-style distinct colors (no similar shades)
const RAINBOW_COLORS = [
  "3B82F6", // Blue
  "10B981", // Green
  "F59E0B", // Orange/Amber
  "06B6D4", // Cyan/Teal
  "8B5CF6", // Purple
  "EF4444", // Red
  "84CC16", // Lime/Yellow-Green
  "EC4899", // Pink
  "14B8A6", // Teal
  "F97316"  // Orange
];

let lastColorIndex = -1;

function randColor() {
  // Get a random color that is different from the last one
  let newIndex;
  do {
    newIndex = Math.floor(Math.random() * RAINBOW_COLORS.length);
  } while (newIndex === lastColorIndex && RAINBOW_COLORS.length > 1);

  lastColorIndex = newIndex;
  return RAINBOW_COLORS[newIndex];
}

// Export for global access
window.resetColors = resetColors;
window.randColor = randColor;

// ========== TEMPLATE REGISTRY ==========
window.AI_TEMPLATES = {};

/************************************************************
 * TEMPLATE 1: TITLE SLIDE
 * Hero layout with bold title and accent elements
 ************************************************************/
AI_TEMPLATES.HERO_IMAGE_TITLE = {
  name: "Title Hero",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    // Background with image or solid color
    if (assets?.bg) {
      s.addImage({ path: assets.bg.path, x: 0, y: 0, w: 10, h: 5.625 });
      // Dark overlay
      s.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 5.625,
        fill: { color: "000000", transparency: 45 }
      });
    } else {
      // Gradient-style split
      s.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 4,
        fill: { color: p.primary }
      });
      s.addShape(pptx.ShapeType.rect, {
        x: 0, y: 4, w: 10, h: 1.625,
        fill: { color: p.dark }
      });
    }

    // Center title
    s.addText(slide.title || "Presentation", {
      x: 0.5, y: 2, w: 9, h: 1.2,
      fontSize: 44,
      bold: true,
      align: "center",
      color: "FFFFFF"
    });

    // Subtitle line
    if (slide.items?.[0]) {
      s.addText(slide.items[0], {
        x: 1.5, y: 3.3, w: 7, h: 0.6,
        fontSize: 18,
        align: "center",
        color: p.accent
      });
    }

    // Bottom accent line
    s.addShape(pptx.ShapeType.rect, {
      x: 4, y: 4.2, w: 2, h: 0.08,
      fill: { color: p.accent }
    });
  }
};

/************************************************************
 * TEMPLATE 2: FEATURE CARDS
 * Left title with right-side gray rounded cards with icons
 * Based on reference image 4 (Feature in LMS portal)
 ************************************************************/
AI_TEMPLATES.LEFT_TITLE_FEATURE_CARDS = {
  name: "Feature Cards",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();

    s.background = { fill: "FFFFFF" };

    // Title on left (bold, black)
    s.addText(slide.title || "", {
      x: 0.4, y: 0.8, w: 3.2, h: 1.8,
      fontSize: 28,
      bold: true,
      color: "1F2937"
    });

    // Feature cards on right (gray background, rounded)
    const items = (slide.items || []).slice(0, 4);
    let y = 0.5;

    items.forEach((text, i) => {
      // Gray rounded card
      s.addShape(pptx.ShapeType.roundRect, {
        x: 3.8, y: y, w: 5.9, h: 1.1,
        fill: { color: "F3F4F6" }
      });

      // Icon placeholder (orange/gray icon)
      if (assets?.icons?.[i]) {
        s.addImage({
          path: assets.icons[i].path,
          x: 4.1, y: y + 0.25, w: 0.6, h: 0.6
        });
      } else {
        // Default icon placeholder
        s.addShape(pptx.ShapeType.ellipse, {
          x: 4.1, y: y + 0.25, w: 0.6, h: 0.6,
          fill: { color: "F59E0B" }
        });
      }

      // Text content
      s.addText(text, {
        x: 5.0, y: y + 0.3, w: 4.5, h: 0.5,
        fontSize: 14,
        color: "374151"
      });

      y += 1.25;
    });
  }
};

/************************************************************
 * TEMPLATE 3: PROCESS STEPS
 * Horizontal stacked bars with dynamic colors
 * Based on reference image 2 (LMS Application Process)
 ************************************************************/
AI_TEMPLATES.PROCESS_VERTICAL_STACK = {
  name: "Process Steps",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();

    s.background = { fill: "FFFFFF" };

    // Title on left (bold, black)
    s.addText(slide.title || "", {
      x: 0.4, y: 0.6, w: 3, h: 1.5,
      fontSize: 26,
      bold: true,
      color: "1F2937"
    });

    // Left accent bar with random color
    s.addShape(pptx.ShapeType.rect, {
      x: 0.35, y: 0.5, w: 0.08, h: 1.8,
      fill: { color: randColor() }
    });

    const steps = (slide.items || []).slice(0, 6);
    let y = 0.5;

    steps.forEach((step, i) => {
      // Use random color for each bar
      const color = randColor();

      // Colored bar (full width on right)
      s.addShape(pptx.ShapeType.roundRect, {
        x: 3.5, y: y, w: 6.2, h: 0.75,
        fill: { color: color }
      });

      // Step text (white, inside bar)
      s.addText(step, {
        x: 3.7, y: y + 0.18, w: 5.8, h: 0.45,
        fontSize: 15,
        color: "FFFFFF"
      });

      y += 0.85;
    });
  }
};

/************************************************************
 * TEMPLATE 4: IMAGE + BULLETS
 * Left image, right bullet list
 ************************************************************/
AI_TEMPLATES.IMAGE_LEFT_BULLETS_RIGHT = {
  name: "Image + Bullets",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    // Left side: image or colored block
    if (assets?.bg) {
      s.addImage({
        path: assets.bg.path,
        x: 0, y: 0, w: 4.5, h: 5.625
      });
    } else {
      s.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 4.5, h: 5.625,
        fill: { color: p.primary }
      });
      // Decorative circle
      s.addShape(pptx.ShapeType.ellipse, {
        x: 1.25, y: 1.8, w: 2, h: 2,
        fill: { color: p.secondary, transparency: 30 }
      });
    }

    // Right white area
    s.addShape(pptx.ShapeType.rect, {
      x: 4.5, y: 0, w: 5.5, h: 5.625,
      fill: { color: "FFFFFF" }
    });

    // Title
    s.addText(slide.title || "", {
      x: 4.8, y: 0.5, w: 5, h: 0.8,
      fontSize: 24,
      bold: true,
      color: p.dark
    });

    // Accent line
    s.addShape(pptx.ShapeType.rect, {
      x: 4.8, y: 1.25, w: 1, h: 0.05,
      fill: { color: p.primary }
    });

    // Bullet points
    const items = (slide.items || []).slice(0, 6);
    let y = 1.5;

    items.forEach((item) => {
      // Bullet dot
      s.addShape(pptx.ShapeType.ellipse, {
        x: 4.8, y: y + 0.15, w: 0.15, h: 0.15,
        fill: { color: p.primary }
      });
      // Text
      s.addText(item, {
        x: 5.1, y: y, w: 4.5, h: 0.55,
        fontSize: 12,
        color: "374151"
      });
      y += 0.6;
    });
  }
};

/************************************************************
 * TEMPLATE 5: GRID BOXES
 * 3-column colorful grid with title on left (dynamic colors)
 * Based on reference image 1 (Step 2 - Upload or Require Details)
 ************************************************************/
AI_TEMPLATES.GRID_REQUIREMENTS = {
  name: "Grid Boxes",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();

    s.background = { fill: "FFFFFF" };

    // Title on left (bold, black with line)
    s.addText(slide.title || "", {
      x: 0.4, y: 0.6, w: 3, h: 1.8,
      fontSize: 28,
      bold: true,
      color: "1F2937"
    });

    // Left accent bar with random color
    s.addShape(pptx.ShapeType.rect, {
      x: 0.35, y: 0.5, w: 0.08, h: 2,
      fill: { color: randColor() }
    });

    const items = (slide.items || []).slice(0, 12);
    const cols = 3;
    const boxW = 1.85;
    const boxH = 1.0;
    const gapX = 0.12;
    const gapY = 0.12;
    const startX = 3.6;
    const startY = 0.5;

    items.forEach((text, i) => {
      const row = Math.floor(i / cols);
      const col = i % cols;
      const x = startX + col * (boxW + gapX);
      const y = startY + row * (boxH + gapY);
      // Use random color for each box
      const color = randColor();

      // Colored box
      s.addShape(pptx.ShapeType.roundRect, {
        x: x, y: y, w: boxW, h: boxH,
        fill: { color: color }
      });

      // Text (white, centered)
      s.addText(text, {
        x: x + 0.1, y: y + 0.25, w: boxW - 0.2, h: boxH - 0.4,
        fontSize: 10,
        align: "center",
        color: "FFFFFF"
      });
    });
  }
};

/************************************************************
 * TEMPLATE 6: FULL BLEED IMAGE / PARAGRAPH CONTENT
 * Full background with text overlay - perfect for paragraphs
 ************************************************************/
AI_TEMPLATES.FULL_BG_TEXT_OVERLAY = {
  name: "Full Bleed / Paragraph",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    // White background for clean paragraph display
    s.background = { fill: "FFFFFF" };

    // Left accent bar
    s.addShape(pptx.ShapeType.rect, {
      x: 0.3, y: 0.4, w: 0.08, h: 1.5,
      fill: { color: p.primary }
    });

    // Title at top left
    s.addText(slide.title || "", {
      x: 0.5, y: 0.5, w: 5, h: 1,
      fontSize: 32,
      bold: true,
      color: "1F2937"
    });

    // Content handling - paragraph style vs bullet style
    if (slide.items?.length) {
      const items = slide.items;

      // Check if this is paragraph content (1-2 long items) or bullet content
      const isParagraph = items.length <= 2 && items.some(i => i.length > 80);

      if (isParagraph) {
        // Display as paragraph text in a colored box
        const fullText = items.join(" ");

        // Colored text box on right side
        s.addShape(pptx.ShapeType.roundRect, {
          x: 3.5, y: 0.4, w: 6, h: 2.5,
          fill: { color: p.primary }
        });

        s.addText(fullText, {
          x: 3.7, y: 0.6, w: 5.6, h: 2.2,
          fontSize: 14,
          color: "FFFFFF",
          valign: "top"
        });
      } else {
        // Display as bullet points for lists
        let y = 1.8;
        const maxItems = items.slice(0, 6);

        maxItems.forEach((item) => {
          // Bullet dot
          s.addShape(pptx.ShapeType.ellipse, {
            x: 0.5, y: y + 0.12, w: 0.12, h: 0.12,
            fill: { color: p.primary }
          });

          s.addText(item, {
            x: 0.75, y: y, w: 8.5, h: 0.5,
            fontSize: 13,
            color: "374151"
          });
          y += 0.55;
        });
      }
    }

    // Add decorative image on bottom right if available
    if (assets?.bg) {
      s.addImage({
        path: assets.bg.path,
        x: 5.5, y: 3.2, w: 4, h: 2.2,
        rounding: true
      });
    }
  }
};

/************************************************************
 * TEMPLATE 7: VALUE CHAIN FLOW
 * Simple horizontal flow with connected boxes (dynamic colors)
 ************************************************************/
AI_TEMPLATES.VALUE_CHAIN_FLOW = {
  name: "Value Chain",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();

    s.background = { fill: "FFFFFF" };

    // Title
    s.addText(slide.title || "", {
      x: 0.5, y: 0.3, w: 9, h: 0.6,
      fontSize: 26,
      bold: true,
      color: "1F2937"
    });

    // Accent line with random color
    s.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 0.85, w: 1.5, h: 0.06,
      fill: { color: randColor() }
    });

    const items = (slide.items || []).slice(0, 5);
    const boxW = 1.7;
    const gapX = 0.15;
    const startX = 0.5;
    const boxY = 1.8;

    // Draw boxes with arrows
    items.forEach((text, i) => {
      const x = startX + i * (boxW + gapX);
      // Use random color for each box
      const color = randColor();

      // Box
      s.addShape(pptx.ShapeType.roundRect, {
        x: x, y: boxY, w: boxW, h: 2.8,
        fill: { color: color }
      });

      // Number
      s.addText(String(i + 1), {
        x: x, y: boxY + 0.1, w: boxW, h: 0.5,
        fontSize: 22,
        bold: true,
        align: "center",
        color: "FFFFFF"
      });

      // Text
      s.addText(text, {
        x: x + 0.1, y: boxY + 0.7, w: boxW - 0.2, h: 2,
        fontSize: 10,
        align: "center",
        color: "FFFFFF"
      });

      // Arrow (except for last item)
      if (i < items.length - 1) {
        s.addText("→", {
          x: x + boxW - 0.1, y: boxY + 1.2, w: 0.4, h: 0.4,
          fontSize: 16,
          color: "6B7280"
        });
      }
    });
  }
};

/************************************************************
 * TEMPLATE 8: TWO COLUMNS WITH IMAGE
 * Clean layout with bullets left, image right
 ************************************************************/
AI_TEMPLATES.TEXT_LEFT_IMAGE_RIGHT = {
  name: "Two Columns",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    s.background = { fill: "FFFFFF" };

    // Title
    s.addText(slide.title || "", {
      x: 0.5, y: 0.3, w: 9, h: 0.7,
      fontSize: 26,
      bold: true,
      color: "1F2937"
    });

    // Accent underline
    s.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 0.95, w: 1.2, h: 0.06,
      fill: { color: "F59E0B" }
    });

    // Bullet points on left
    const items = (slide.items || []).slice(0, 6);
    let y = 1.3;

    items.forEach((item) => {
      s.addShape(pptx.ShapeType.ellipse, {
        x: 0.5, y: y + 0.15, w: 0.15, h: 0.15,
        fill: { color: p.primary }
      });
      s.addText(item, {
        x: 0.8, y: y, w: 4.2, h: 0.55,
        fontSize: 13,
        color: "374151"
      });
      y += 0.65;
    });

    // Image on right with rounded corners effect
    if (assets?.bg) {
      s.addImage({
        path: assets.bg.path,
        x: 5.2, y: 1.1, w: 4.3, h: 4,
        rounding: true
      });
    } else {
      // Placeholder with gradient
      s.addShape(pptx.ShapeType.roundRect, {
        x: 5.2, y: 1.1, w: 4.3, h: 4,
        fill: { color: p.light },
        line: { color: p.primary, width: 2 }
      });
    }
  }
};

/************************************************************
 * TEMPLATE 9: SOLUTION CARDS
 * 3 clean cards with icons and title only
 ************************************************************/
AI_TEMPLATES.SOLUTION_CARDS_GRID = {
  name: "Solution Cards Grid",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    s.background = { fill: "FFFFFF" };

    // Title Centered
    s.addText(slide.title || "", {
      x: 0.5, y: 0.3, w: 9, h: 0.7,
      fontSize: 26,
      bold: true,
      align: "center",
      color: p.dark
    });

    // Accent underline
    s.addShape(pptx.ShapeType.rect, {
      x: 4.3, y: 0.95, w: 1.4, h: 0.06,
      fill: { color: p.primary }
    });

    const items = (slide.items || []).slice(0, 3);
    const cardW = 2.9;
    const startX = 0.65;
    const gap = 3.05;

    items.forEach((text, i) => {
      const x = startX + i * gap;
      const cardPalette = COLOR_PALETTES[(paletteIndex + i) % COLOR_PALETTES.length];

      // Clean card background (no border lines)
      s.addShape(pptx.ShapeType.roundRect, {
        x: x, y: 1.2, w: cardW, h: 3.9,
        fill: { color: cardPalette.light }
      });

      // Icon circle
      s.addShape(pptx.ShapeType.ellipse, {
        x: x + 0.95, y: 1.5, w: 1, h: 1,
        fill: { color: cardPalette.primary }
      });

      // Use actual icon if available
      if (assets?.icons?.[i]) {
        s.addImage({
          path: assets.icons[i].path,
          x: x + 1.1, y: 1.65, w: 0.7, h: 0.7
        });
      }

      // Title only (no duplicate description)
      s.addText(text, {
        x: x + 0.15, y: 2.7, w: cardW - 0.3, h: 2.2,
        fontSize: 12,
        align: "center",
        color: cardPalette.dark
      });
    });
  }
};

/************************************************************
 * TEMPLATE 10: BLOG CARDS LAYOUT
 * Simple 3-column content cards
 ************************************************************/
AI_TEMPLATES.BLOG_CARDS_LAYOUT = {
  name: "Content Cards",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    s.background = { fill: "F8FAFC" };

    // Title
    s.addText(slide.title || "", {
      x: 0.5, y: 0.3, w: 9, h: 0.6,
      fontSize: 26,
      bold: true,
      color: "1F2937"
    });

    // Accent line
    s.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 0.85, w: 1.5, h: 0.06,
      fill: { color: "F59E0B" }
    });

    const items = (slide.items || []).slice(0, 6);
    const cols = 3;
    const cardW = 2.95;
    const cardH = 1.9;
    const gapX = 0.2;
    const gapY = 0.2;
    const startX = 0.5;
    const startY = 1.1;
    const cardColors = ["3B82F6", "10B981", "F59E0B", "EC4899", "8B5CF6", "EF4444"];

    items.forEach((text, i) => {
      const row = Math.floor(i / cols);
      const col = i % cols;
      const x = startX + col * (cardW + gapX);
      const y = startY + row * (cardH + gapY);
      const color = cardColors[i % cardColors.length];

      // White card
      s.addShape(pptx.ShapeType.roundRect, {
        x: x, y: y, w: cardW, h: cardH,
        fill: { color: "FFFFFF" },
        line: { color: "E5E7EB", width: 1 }
      });

      // Colored top bar
      s.addShape(pptx.ShapeType.rect, {
        x: x, y: y, w: cardW, h: 0.08,
        fill: { color: color }
      });

      // Number badge
      s.addShape(pptx.ShapeType.ellipse, {
        x: x + 0.15, y: y + 0.2, w: 0.4, h: 0.4,
        fill: { color: color }
      });
      s.addText(String(i + 1), {
        x: x + 0.15, y: y + 0.25, w: 0.4, h: 0.35,
        fontSize: 12,
        bold: true,
        align: "center",
        color: "FFFFFF"
      });

      // Text
      s.addText(text, {
        x: x + 0.15, y: y + 0.7, w: cardW - 0.3, h: cardH - 0.9,
        fontSize: 11,
        color: "374151"
      });
    });
  }
};

/************************************************************
 * TEMPLATE 11: INSTALLATION GALLERY
 * Colorful gradient cards in 2 rows (like reference image 1)
 ************************************************************/
AI_TEMPLATES.INSTALLATION_STEPS_CARDS = {
  name: "Gallery Cards",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();

    s.background = { fill: "F3F4F6" };

    // Title
    s.addText(slide.title || "", {
      x: 0.5, y: 0.2, w: 9, h: 0.6,
      fontSize: 28,
      bold: true,
      align: "center",
      color: "1F2937"
    });

    // Subtitle
    s.addText("Step-by-step process overview", {
      x: 0.5, y: 0.75, w: 9, h: 0.35,
      fontSize: 11,
      align: "center",
      color: "6B7280"
    });

    const items = (slide.items || []).slice(0, 6);
    const cols = 4;
    const cardW = 2.2;
    const cardH = 1.8;
    const gapX = 0.2;
    const gapY = 0.2;
    const startX = 0.4;
    const startY = 1.2;

    // Gradient-like solid colors
    const cardColors = [
      "10B981", // Emerald
      "0EA5E9", // Sky blue  
      "06B6D4", // Cyan
      "A855F7", // Purple gradient end
      "F97316", // Orange
      "3B82F6"  // Blue
    ];

    items.forEach((text, i) => {
      const row = Math.floor(i / cols);
      const col = i % cols;
      const x = startX + col * (cardW + gapX);
      const y = startY + row * (cardH + gapY);
      const color = cardColors[i % cardColors.length];

      // Card with rounded corners
      s.addShape(pptx.ShapeType.roundRect, {
        x: x, y: y, w: cardW, h: cardH,
        fill: { color: color }
      });

      // Large number (watermark style)
      s.addText(String(i + 1), {
        x: x + 0.1, y: y + 0.05, w: 0.6, h: 0.7,
        fontSize: 40,
        bold: true,
        color: "FFFFFF",
        transparency: 40
      });

      // Title (first 2-3 words)
      const words = text.split(" ");
      s.addText(words.slice(0, 3).join(" "), {
        x: x + 0.15, y: y + 0.75, w: cardW - 0.3, h: 0.45,
        fontSize: 12,
        bold: true,
        color: "FFFFFF"
      });

      // Description (remaining words)
      s.addText(words.slice(3).join(" ") || text, {
        x: x + 0.15, y: y + 1.2, w: cardW - 0.3, h: 0.5,
        fontSize: 9,
        color: "FFFFFF",
        transparency: 15
      });
    });
  }
};

/************************************************************
 * TEMPLATE 12: ZIGZAG PROCESS
 * Simple vertical numbered steps
 ************************************************************/
AI_TEMPLATES.ZIGZAG_PROCESS_STEPS = {
  name: "Vertical Steps",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();

    s.background = { fill: "FFFFFF" };

    // Title
    s.addText(slide.title || "", {
      x: 0.5, y: 0.3, w: 9, h: 0.6,
      fontSize: 26,
      bold: true,
      color: "1F2937"
    });

    // Accent line
    s.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 0.85, w: 1.5, h: 0.06,
      fill: { color: "F59E0B" }
    });

    const items = (slide.items || []).slice(0, 5);
    const stepColors = ["3B82F6", "10B981", "F59E0B", "EC4899", "8B5CF6"];
    let y = 1.2;

    items.forEach((text, i) => {
      const color = stepColors[i % stepColors.length];

      // Number circle
      s.addShape(pptx.ShapeType.ellipse, {
        x: 0.5, y: y + 0.1, w: 0.6, h: 0.6,
        fill: { color: color }
      });
      s.addText(String(i + 1), {
        x: 0.5, y: y + 0.17, w: 0.6, h: 0.5,
        fontSize: 16,
        bold: true,
        align: "center",
        color: "FFFFFF"
      });

      // Vertical line (except for last)
      if (i < items.length - 1) {
        s.addShape(pptx.ShapeType.rect, {
          x: 0.77, y: y + 0.7, w: 0.06, h: 0.2,
          fill: { color: "D1D5DB" }
        });
      }

      // Content bar
      s.addShape(pptx.ShapeType.roundRect, {
        x: 1.3, y: y, w: 8.2, h: 0.75,
        fill: { color: "F9FAFB" },
        line: { color: "E5E7EB", width: 1 }
      });

      // Text
      s.addText(text, {
        x: 1.5, y: y + 0.2, w: 7.8, h: 0.45,
        fontSize: 13,
        color: "374151"
      });

      y += 0.9;
    });
  }
};

/************************************************************
 * TEMPLATE 13: TEXT LEFT + IMAGE RIGHT
 * Same style as Template 4 but reversed (text left, image right)
 ************************************************************/
AI_TEMPLATES.TEXT_LEFT_IMAGE_RIGHT = {
  name: "Text Left + Image Right",

  build: function (pptx, slide, assets) {
    const s = pptx.addSlide();
    const p = getNextPalette();

    // Left white area
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 5, h: 5.625,
      fill: { color: "FFFFFF" }
    });

    // Right side: image or colored block
    if (assets?.bg) {
      s.addImage({
        path: assets.bg.path,
        x: 5, y: 0, w: 5, h: 5.625
      });
    } else {
      s.addShape(pptx.ShapeType.rect, {
        x: 5, y: 0, w: 5, h: 5.625,
        fill: { color: p.primary }
      });
      // Decorative circle
      s.addShape(pptx.ShapeType.ellipse, {
        x: 6.5, y: 1.8, w: 2, h: 2,
        fill: { color: p.secondary, transparency: 30 }
      });
    }

    // Title
    s.addText(slide.title || "", {
      x: 0.5, y: 0.5, w: 4.2, h: 0.8,
      fontSize: 24,
      bold: true,
      color: p.dark
    });

    // Accent line
    s.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.25, w: 1, h: 0.05,
      fill: { color: p.primary }
    });

    // Bullet points
    const items = (slide.items || []).slice(0, 8);
    let y = 1.5;

    items.forEach((item) => {
      // Bullet dot
      s.addShape(pptx.ShapeType.ellipse, {
        x: 0.5, y: y + 0.15, w: 0.15, h: 0.15,
        fill: { color: p.primary }
      });

      // Text
      s.addText(item, {
        x: 0.8, y: y, w: 3.9, h: 0.5,
        fontSize: 12,
        color: "374151"
      });

      y += 0.5;
    });
  }
};

console.log("Prabhas PPT Maker - Templates loaded ✅");

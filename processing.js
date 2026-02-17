/**
 * Prabhas PPT Maker - Frontend Processing Engine v3.0
 * Pure JavaScript text processor for AI-generated content
 * No backend required - works entirely in browser
 */

// ============== HELPER FUNCTIONS ==============

function cleanMarkdown(text) {
    if (!text) return "";
    // Remove bold markers
    text = text.replace(/\*\*([^*]*)\*\*/g, '$1');
    text = text.replace(/__([^_]*)__/g, '$1');
    // Remove italic markers
    text = text.replace(/\*([^*]+)\*/g, '$1');
    text = text.replace(/_([^_]+)_/g, '$1');
    // Remove markdown headers
    text = text.replace(/^#{1,6}\s*/gm, '');
    // Remove code blocks
    text = text.replace(/`([^`]+)`/g, '$1');
    // Remove stray asterisks
    text = text.replace(/^\*+\s*/g, '');
    text = text.replace(/\s*\*+$/g, '');
    // Clean whitespace
    text = text.replace(/\s+/g, ' ');
    return text.trim();
}

function isSkipLine(line) {
    const lineLower = line.trim().toLowerCase();

    // Empty or separator
    if (!lineLower || lineLower === '---' || lineLower.startsWith('```')) {
        return true;
    }

    // Pure labels
    const labels = ['title:', 'content:', 'overview:', 'focus:', 'subtitle:',
        'definition:', 'causes:', 'effects:', 'conclusion:',
        'presented by:', 'date:', 'your name'];
    if (labels.includes(lineLower.replace(/:$/, '') + ':')) {
        return true;
    }

    // AI conversation lines
    const skipPhrases = ['would you like', 'tell me your', 'if you want',
        "here's", "here is", "i can also", "let me know",
        "sure", "sureee", "ready-to-copy", "presentation-friendly",
        "slide 2,3,4", "slide 2, 3, 4"];
    return skipPhrases.some(p => lineLower.includes(p));
}

function extractLabelContent(line) {
    // **Label:** content
    let match = line.match(/\*\*([^*:]+):\*\*\s*(.+)/);
    if (match) {
        const label = match[1].trim().toLowerCase();
        const content = cleanMarkdown(match[2]);
        if (['title', 'main title'].includes(label)) return { type: 'title', content };
        if (['overview', 'focus', 'content', 'description', 'definition'].includes(label)) return { type: 'overview', content };
        if (label === 'conclusion') return { type: 'conclusion', content };
    }

    // **Label**: content
    match = line.match(/\*\*([^*]+)\*\*:\s*(.+)/);
    if (match) {
        const label = match[1].trim().toLowerCase();
        const content = cleanMarkdown(match[2]);
        if (['title', 'main title'].includes(label)) return { type: 'title', content };
        if (['overview', 'focus', 'content'].includes(label)) return { type: 'overview', content };
    }

    // Label: content
    match = line.match(/^([A-Za-z\s]+):\s*(.+)/);
    if (match) {
        const label = match[1].trim().toLowerCase();
        const content = cleanMarkdown(match[2]);
        if (['title', 'main title'].includes(label)) return { type: 'title', content };
        if (['overview', 'focus', 'content', 'definition'].includes(label)) return { type: 'overview', content };
        if (label === 'conclusion') return { type: 'conclusion', content };
    }

    return null;
}

function parseSlideContent(contentText) {
    const lines = contentText.split('\n');
    let title = null;
    let items = [];
    let overview = null;
    let conclusion = null;

    for (let line of lines) {
        line = line.trim();

        if (isSkipLine(line)) continue;

        // Check labeled content
        const labelData = extractLabelContent(line);
        if (labelData) {
            if (labelData.type === 'title') {
                title = labelData.content;
                continue;
            } else if (labelData.type === 'overview') {
                overview = labelData.content;
                continue;
            } else if (labelData.type === 'conclusion') {
                conclusion = labelData.content;
                continue;
            }
        }

        // Bullet points
        let bulletMatch = line.match(/^[-•*►▪◦]\s*(.+)/);
        if (bulletMatch) {
            const item = cleanMarkdown(bulletMatch[1]);
            if (item && item.length > 3 && !isSkipLine(item)) {
                items.push(item);
            }
            continue;
        }

        // Numbered items
        let numberMatch = line.match(/^\d+[.)]\s*(.+)/);
        if (numberMatch) {
            const item = cleanMarkdown(numberMatch[1]);
            if (item && item.length > 3 && !isSkipLine(item)) {
                items.push(item);
            }
            continue;
        }

        // Plain text
        const cleanLine = cleanMarkdown(line);
        if (cleanLine && cleanLine.length > 5 && !isSkipLine(cleanLine)) {
            if (!title) {
                title = cleanLine;
            } else if (cleanLine.length > 20) {
                items.push(cleanLine);
            }
        }
    }

    // Add overview and conclusion
    if (overview && overview.length > 10) {
        items.unshift(overview);
    }
    if (conclusion && conclusion.length > 10) {
        items.push(conclusion);
    }

    return { title, items: items.slice(0, 10) };
}

function extractKeywords(title, items) {
    const text = ((title || '') + ' ' + items.slice(0, 3).join(' ')).toLowerCase();

    const stopWords = new Set([
        'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been',
        'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would',
        'could', 'should', 'may', 'might', 'must', 'shall', 'can',
        'and', 'or', 'but', 'if', 'then', 'else', 'when', 'where',
        'how', 'what', 'which', 'who', 'whom', 'this', 'that', 'these',
        'those', 'am', 'it', 'its', 'of', 'to', 'in', 'for', 'on',
        'with', 'at', 'by', 'from', 'as', 'into', 'through', 'during',
        'before', 'after', 'above', 'below', 'between', 'under', 'again',
        'further', 'once', 'here', 'there', 'all', 'each', 'few', 'more',
        'most', 'other', 'some', 'such', 'no', 'nor', 'not', 'only',
        'own', 'same', 'so', 'than', 'too', 'very', 'just', 'also'
    ]);

    const words = text.match(/\b[a-z]{4,}\b/g) || [];
    const keywords = [];
    const seen = new Set();

    for (const word of words) {
        if (!stopWords.has(word) && !seen.has(word)) {
            keywords.push(word);
            seen.add(word);
            if (keywords.length >= 8) break;
        }
    }

    return keywords;
}

function detectContentType(title, itemCount, slideIndex) {
    const titleLower = (title || '').toLowerCase();

    if (slideIndex === 0) return 'title';

    if (['thank', 'question', 'q&a', 'contact'].some(p => titleLower.includes(p))) {
        return 'thankyou';
    }
    if (['conclusion', 'summary', 'outlook', 'future', 'final'].some(p => titleLower.includes(p))) {
        return 'conclusion';
    }
    if (['introduction', 'overview', 'landscape', 'what is'].some(p => titleLower.includes(p))) {
        return 'overview';
    }
    if (['type', 'kind', 'category', 'classification'].some(p => titleLower.includes(p))) {
        return 'grid';
    }
    if (['step', 'process', 'how to', 'method'].some(p => titleLower.includes(p))) {
        return 'process';
    }
    if (['skill', 'feature', 'benefit', 'require', 'tip'].some(p => titleLower.includes(p))) {
        return 'features';
    }
    if (['company', 'recruiter', 'hire', 'top', 'major'].some(p => titleLower.includes(p))) {
        return 'grid';
    }
    if (['role', 'job', 'career', 'opportunit', 'position'].some(p => titleLower.includes(p))) {
        return 'bullets';
    }

    // Default based on item count
    if (itemCount === 0) return 'paragraph';
    if (itemCount <= 3) return 'features';
    if (itemCount <= 5) return 'bullets';
    return 'grid';
}

// ============== TEMPLATE SELECTION WITH RANDOMIZATION ==============

// Template pools for each content type - adds variety!
const TEMPLATE_POOLS = {
    title: ['HERO_IMAGE_TITLE'],
    thankyou: ['FULL_BG_TEXT_OVERLAY', 'TEXT_LEFT_IMAGE_RIGHT'],
    conclusion: ['TEXT_LEFT_IMAGE_RIGHT', 'FULL_BG_TEXT_OVERLAY', 'IMAGE_LEFT_BULLETS_RIGHT'],
    overview: ['FULL_BG_TEXT_OVERLAY', 'TEXT_LEFT_IMAGE_RIGHT', 'IMAGE_LEFT_BULLETS_RIGHT'],
    paragraph: ['FULL_BG_TEXT_OVERLAY', 'TEXT_LEFT_IMAGE_RIGHT'],
    process: ['ZIGZAG_PROCESS_STEPS', 'PROCESS_VERTICAL_STACK', 'VALUE_CHAIN_FLOW'],
    features: ['SOLUTION_CARDS_GRID', 'LEFT_TITLE_FEATURE_CARDS', 'BLOG_CARDS_LAYOUT'],
    grid: ['GRID_REQUIREMENTS', 'INSTALLATION_STEPS_CARDS', 'BLOG_CARDS_LAYOUT'],
    bullets: ['IMAGE_LEFT_BULLETS_RIGHT', 'TEXT_LEFT_IMAGE_RIGHT', 'PROCESS_VERTICAL_STACK', 'LEFT_TITLE_FEATURE_CARDS']
};

function suggestTemplate(contentType, itemCount) {
    const pool = TEMPLATE_POOLS[contentType] || TEMPLATE_POOLS.bullets;

    // Random selection from pool for variety!
    const randomIndex = Math.floor(Math.random() * pool.length);
    return pool[randomIndex];
}

// ============== MAIN PROCESSING FUNCTION ==============

function splitIntoSlides(text) {
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

    // Pattern for slide markers
    const slidePattern = /(?:^|\n)(?:#{1,3}\s*)?(?:\*\*)?Slide\s*(\d+)\s*[:\–\-]?\s*(?:\*\*)?\s*([^\n]*)/gi;

    const matches = [];
    let match;
    while ((match = slidePattern.exec(text)) !== null) {
        matches.push({
            index: match.index,
            end: slidePattern.lastIndex,
            number: match[1],
            header: cleanMarkdown(match[2] || '')
        });
    }

    if (matches.length === 0) return null;

    const slidesRaw = [];
    for (let i = 0; i < matches.length; i++) {
        const start = matches[i].end;
        const end = (i + 1 < matches.length) ? matches[i + 1].index : text.length;
        const content = text.substring(start, end).trim();

        slidesRaw.push({
            number: matches[i].number,
            header: matches[i].header,
            content: content
        });
    }

    return slidesRaw;
}

/**
 * Main processing function - call this from index.html
 * @param {string} text - Raw AI-generated content
 * @returns {Object} - { slides: [...], error: string|null }
 */
function processContent(text) {
    text = text.trim();
    if (!text) {
        return { slides: [], error: "No content provided" };
    }

    console.log("🎯 Processing content...", text.length, "chars");

    // Split by slide markers
    const slidesRaw = splitIntoSlides(text);

    if (!slidesRaw || slidesRaw.length < 1) {
        console.log("❌ No slide markers found");
        return {
            slides: [],
            error: "No slide structure found. Please paste AI-generated content with 'Slide 1:', 'Slide 2:', etc."
        };
    }

    console.log("📊 Found", slidesRaw.length, "slides");

    const slides = [];
    for (let i = 0; i < slidesRaw.length && i < 16; i++) {
        const slideData = slidesRaw[i];

        // Parse content
        const { title: parsedTitle, items } = parseSlideContent(slideData.content);

        // Use header as title if no title in content
        let title = parsedTitle || slideData.header || `Slide ${i + 1}`;

        // Detect type and suggest template (with randomization!)
        const contentType = detectContentType(title, items.length, slides.length);
        const suggestedTemplate = suggestTemplate(contentType, items.length);

        // Extract keywords
        const keywords = extractKeywords(title, items);

        slides.push({
            title: title.substring(0, 80),
            items: items,
            keywords: keywords,
            content_type: contentType,
            suggested_template: suggestedTemplate
        });

        console.log(`  ✅ Slide ${slides.length}: '${title.substring(0, 40)}' | ${items.length} items | ${suggestedTemplate}`);
    }

    console.log("📈 Total:", slides.length, "slides generated");

    return { slides: slides, error: null };
}

// Export for use in index.html
window.processContent = processContent;

console.log("✅ Prabhas PPT - Processing Engine v3.0 Loaded (Frontend-Only)");

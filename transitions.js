/**
 * Prabhas PPT Maker - Slide Transitions Engine
 * Injects real PowerPoint transitions into generated PPTX files
 * Uses JSZip to modify the PPTX (ZIP) archive after PptxGenJS generates it
 * 
 * Supported transitions:
 *   - Morph (requires PowerPoint 2019+ / Office 365)
 *   - Push (universal compatibility)
 * 
 * Usage: Call window.applyTransitions(blob) with a PPTX Blob
 *        Returns a new Blob with transitions injected
 */

(function () {
    'use strict';

    // Transition XML templates for PowerPoint slides
    // These get injected into each slide's XML in the PPTX archive
    const TRANSITIONS = {
        // Morph transition (Office 365 / PowerPoint 2019+)
        morph: '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p159"><p:transition spd="slow" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>',

        // Push transition (universal)
        push: '<p:transition spd="med"><p:push dir="l"/></p:transition>',

        // Cover transition (universal) 
        cover: '<p:transition spd="med"><p:cover dir="l"/></p:transition>',

        // Fade transition (universal fallback)
        fade: '<p:transition spd="slow"><p:fade/></p:transition>',

        // Wipe transition 
        wipe: '<p:transition spd="med"><p:wipe dir="d"/></p:transition>'
    };

    // Alternating pattern: morph → push → morph → push...
    const TRANSITION_SEQUENCE = ['morph', 'push'];

    /**
     * Apply transitions to a PPTX blob
     * @param {Blob} pptxBlob - The original PPTX file as a Blob
     * @returns {Promise<Blob>} - Modified PPTX blob with transitions
     */
    async function applyTransitions(pptxBlob) {
        // Dynamically load JSZip if not present
        if (typeof JSZip === 'undefined') {
            await loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
        }

        try {
            const zip = await JSZip.loadAsync(pptxBlob);
            const slideFiles = [];

            // Find all slide XML files
            zip.forEach(function (relativePath) {
                if (/^ppt\/slides\/slide\d+\.xml$/.test(relativePath)) {
                    slideFiles.push(relativePath);
                }
            });

            // Sort slides numerically
            slideFiles.sort(function (a, b) {
                var numA = parseInt(a.match(/slide(\d+)/)[1]);
                var numB = parseInt(b.match(/slide(\d+)/)[1]);
                return numA - numB;
            });

            console.log('[Transitions] Found ' + slideFiles.length + ' slides');

            // Inject transitions into each slide (skip slide 1)
            for (var i = 1; i < slideFiles.length; i++) {
                var slideXml = await zip.file(slideFiles[i]).async('string');
                var transType = TRANSITION_SEQUENCE[i % TRANSITION_SEQUENCE.length];
                var transXml = TRANSITIONS[transType];

                // Insert transition XML before closing </p:sld> tag
                if (slideXml.indexOf('<p:transition') === -1) {
                    slideXml = slideXml.replace('</p:sld>', transXml + '</p:sld>');
                    zip.file(slideFiles[i], slideXml);
                    console.log('[Transitions] Slide ' + (i + 1) + ': ' + transType);
                }
            }

            // Generate modified PPTX
            var modifiedBlob = await zip.generateAsync({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 }
            });

            console.log('[Transitions] ✅ All transitions applied');
            return modifiedBlob;

        } catch (err) {
            console.error('[Transitions] Error:', err);
            // Return original blob if anything fails
            return pptxBlob;
        }
    }

    /**
     * Helper: Download a Blob as a file
     */
    function downloadBlob(blob, filename) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    /**
     * Helper: Dynamically load a script
     */
    function loadScript(src) {
        return new Promise(function (resolve, reject) {
            if (document.querySelector('script[src="' + src + '"]')) {
                resolve();
                return;
            }
            var s = document.createElement('script');
            s.src = src;
            s.onload = resolve;
            s.onerror = reject;
            document.head.appendChild(s);
        });
    }

    // Export globally
    window.applyTransitions = applyTransitions;
    window.downloadBlobFile = downloadBlob;

    console.log('✅ Prabhas PPT - Transitions Engine Loaded');
})();

// Caseware URL patterns:
// Engagement-level: https://<host>/<tenant>/e/eng/<engagementId>/...
// With document:    ...#/efinancials/<documentId>
const CW_URL_PATTERN = /https?:\/\/([^/]+)\/([^/]+)\/e\/eng\/([^/]+)/;
const CW_DOC_PATTERN = /#\/(?:efinancials|checklist)\/([^/?\s]+)/;

const urlInput = document.getElementById('cwUrl');
const parsedInfo = document.getElementById('parsedInfo');
const generateBtn = document.getElementById('generateBtn');
const statusEl = document.getElementById('status');
const errorEl = document.getElementById('error');
const successEl = document.getElementById('success');

// Live URL parsing feedback
urlInput.addEventListener('input', function () {
    const match = this.value.trim().match(CW_URL_PATTERN);
    if (match) {
        const docMatch = this.value.trim().match(CW_DOC_PATTERN);
        let info = 'Tenant: ' + match[2] +
            '  |  Engagement: ' + match[3].slice(0, 12) + '\u2026';
        if (docMatch) {
            info += '  |  Document: ' + docMatch[1].slice(0, 12) + '\u2026';
            info += '  \u2014  will extract this checklist only';
        } else {
            info += '  \u2014  will extract all checklists';
        }
        parsedInfo.textContent = info;
        parsedInfo.classList.add('parsed-success');
        urlInput.classList.remove('input-error');
    } else if (this.value.trim()) {
        parsedInfo.textContent = 'URL not recognized \u2014 expected a Caseware engagement URL';
        parsedInfo.classList.remove('parsed-success');
    } else {
        parsedInfo.textContent = 'Paste a Caseware engagement URL to extract all checklists, or a document URL for a single checklist';
        parsedInfo.classList.remove('parsed-success');
    }
});

// Generate report
generateBtn.addEventListener('click', async function () {
    clearMessages();

    const url = urlInput.value.trim();
    if (!url) {
        showError('Please paste a Caseware URL.');
        urlInput.classList.add('input-error');
        urlInput.focus();
        return;
    }

    const match = url.match(CW_URL_PATTERN);
    if (!match) {
        showError('Invalid URL format. Expected a Caseware engagement URL (https://<host>/<tenant>/e/eng/<id>/...)');
        urlInput.classList.add('input-error');
        urlInput.focus();
        return;
    }

    const templateName = document.getElementById('templateName').value.trim();

    setLoading(true);

    try {
        const response = await fetch('/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url, templateName: templateName || 'Report' }),
        });

        if (!response.ok) {
            let errMsg = 'An error occurred while generating the report.';
            try {
                const err = await response.json();
                errMsg = err.error || errMsg;
            } catch (_) { /* response wasn't JSON */ }
            throw new Error(errMsg);
        }

        // Download the file
        const blob = await response.blob();
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;

        // Extract filename from Content-Disposition or use default
        const disposition = response.headers.get('Content-Disposition');
        const filenameMatch = disposition && disposition.match(/filename="?([^"]+)"?/);
        a.download = filenameMatch ? filenameMatch[1] : 'checklists.xlsx';

        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(blobUrl);

        showSuccess('Report downloaded successfully.');
    } catch (e) {
        showError(e.message);
    } finally {
        setLoading(false);
    }
});

// Allow Enter key to submit
urlInput.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') generateBtn.click();
});

function setLoading(loading) {
    generateBtn.disabled = loading;
    statusEl.hidden = !loading;
}

function showError(msg) {
    errorEl.textContent = msg;
    errorEl.hidden = false;
}

function showSuccess(msg) {
    successEl.textContent = msg;
    successEl.hidden = false;
}

function clearMessages() {
    errorEl.hidden = true;
    successEl.hidden = true;
    urlInput.classList.remove('input-error');
}

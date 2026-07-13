// Watches a publish land: polls the session, moves the bar, then hands off to the workbook.
//
// This page deliberately does NOT use a Blazor circuit. Its whole job is to poll one endpoint and
// update three elements; doing that over SignalR would mean a WebSocket, per-user circuit memory,
// and reconnect behaviour through the cloudflared tunnel — all so we can display a percentage.
// Plain fetch has none of that, and still works if WebSockets are blocked anywhere along the path.

(function () {
    const root = document.getElementById('publish');
    if (!root) return;

    const id = root.dataset.session;
    const titleEl = document.getElementById('stage-title');
    const detailEl = document.getElementById('stage-detail');
    const barEl = document.getElementById('bar-fill');
    const errorEl = document.getElementById('stage-error');
    const actionEl = document.getElementById('stage-action');

    const HEADINGS = {
        waiting: 'Waiting for the file…',
        receiving: 'Receiving the file…',
        opening: 'Opening the workbook…',
        ready: 'Ready',
    };

    function size(bytes) {
        if (bytes === null || bytes === undefined) return '';
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / 1048576).toFixed(1) + ' MB';
    }

    function fail(message) {
        titleEl.textContent = 'That workbook couldn’t be opened';
        detailEl.textContent = '';
        barEl.parentElement.style.display = 'none';
        errorEl.textContent = message;
        errorEl.style.display = 'block';
        actionEl.style.display = 'inline-block';
    }

    async function poll() {
        let res;
        try {
            res = await fetch('/api/sessions/' + id, { cache: 'no-store' });
        } catch {
            // A blip on the network is not a failure — the upload is still running server-side.
            return setTimeout(poll, 1000);
        }

        if (res.status === 404) {
            titleEl.textContent = 'This link has expired';
            detailEl.textContent = 'The app that opened this never finished sending the file. Ask it to try again.';
            barEl.parentElement.style.display = 'none';
            actionEl.style.display = 'inline-block';
            return;
        }

        const s = await res.json();

        if (s.stage === 'ready' && s.hash) {
            // replace(), not assign(): the progress page must not sit in the back-stack, or Back
            // from the workbook lands the user on a finished spinner.
            window.location.replace('/view/' + s.hash);
            return;
        }

        if (s.stage === 'failed') return fail(s.error || 'The workbook could not be opened.');

        titleEl.textContent = HEADINGS[s.stage] || '';

        if (s.stage === 'receiving' && s.percent !== null && s.percent !== undefined) {
            barEl.classList.remove('indeterminate');
            barEl.style.width = s.percent + '%';
            detailEl.textContent = s.percent + '% of ' + size(s.totalBytes);
        } else {
            // No declared size, or we're parsing: there is no honest percentage, so show motion
            // rather than invent a number.
            barEl.classList.add('indeterminate');
            barEl.style.width = '';
            detailEl.textContent = s.stage === 'opening'
                ? 'Reading the sheets and applying formatting. Large reports take a few seconds.'
                : 'The app that sent you here is preparing the report.';
        }

        setTimeout(poll, 300);
    }

    poll();
})();

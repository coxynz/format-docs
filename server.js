const express = require('express');
const path = require('path');
const open = require('open');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files from the current directory
app.use(express.static(__dirname));

// Direct route for index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, async () => {
    console.log(`🚀 Format Docs server starting on http://localhost:${PORT}`);
    console.log(`📂 Press Ctrl+C to stop the server`);

    // Attempt to open browser automatically
    try {
        await open(`http://localhost:${PORT}`);
    } catch (e) {
        // Fallback or ignore if 'open' is not available
        console.log(`Please open http://localhost:${PORT} in your browser.`);
    }
});

let dbx;  // Global Dropbox instance
let isExporting = false;  // Flag to control export loop

function sanitizeFilename(filename) {
    return filename.replace(/[<>:"/\\|?*]/g, '');
}

function getRelativePath(entry) {
    // Extract relative path from path_display (e.g., "/folder/sub/file.paper" -> "folder/sub/file.md")
    const path = entry.path_display.replace(/\.paper$/, '.md');
    const parts = path.split('/').filter(part => part !== '');  // Split and clean
    return {
        dir: parts.slice(0, -1).join('/'),  // Folder path (empty if root)
        name: parts[parts.length - 1]  // Filename
    };
}

async function collectPaperFiles() {
    try {
        let result = await dbx.filesListFolder({ path: '', recursive: true });
        let files = [];

        // Process initial batch
        for (let entry of result.result.entries) {
            if (entry['.tag'] === 'file' && entry.name.toLowerCase().endsWith('.paper')) {
                files.push(entry);
            }
        }

        // Pagination
        while (result.result.has_more) {
            result = await dbx.filesListFolderContinue({ cursor: result.result.cursor });
            for (let entry of result.result.entries) {
                if (entry['.tag'] === 'file' && entry.name.toLowerCase().endsWith('.paper')) {
                    files.push(entry);
                }
            }
        }

        return files;
    } catch (error) {
        throw new Error(`Failed to list files: ${error.error_summary || error.message}`);
    }
}

async function exportSingleDoc(entry, zip = null) {
    try {
        const exportResponse = await dbx.filesExport({ path: entry.path_lower, export_format: 'markdown' });
        const contentBlob = exportResponse.result.fileBlob;
        const content = await contentBlob.text();  // Extract Markdown text

        let title = entry.name.replace('.paper', '.md');
        if (!title.endsWith('.md')) title += '.md';
        title = sanitizeFilename(title);

        const { dir } = getRelativePath(entry);

        if (zip) {
            // Add to ZIP with folder structure
            let zipPath = title;  // Default flat, but always preserve since ZIP is for structure
            if (dir) {
                zipPath = `${dir}/${title}`;
            }
            zip.file(zipPath, content);
            updateStatus(`Added to ZIP: ${zipPath}`);
        } else {
            // Individual download: Flattened with  ___  for structure
            let downloadName = title;
            if (dir) {
                const dirParts = dir.split('/').map(sanitizeFilename);
                downloadName = `${dirParts.join(' ___ ')}${dirParts.length > 0 ? ' ___ ' : ''}${title}`;
            }
            downloadName = sanitizeFilename(downloadName);  // Final clean

            const blob = new Blob([content], { type: 'text/markdown' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = downloadName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            updateStatus(`Exported: ${downloadName}`);
        }

        return 1;
    } catch (error) {
        console.error('Export error details:', error);
        if (error.error && error.error.path && error.error.path['.tag'] === 'not_found') {
            updateStatus(`Doc not found: ${entry.path_display}`, 'error');
        } else {
            updateStatus(`Export error for ${entry.path_display}: ${error.error_summary || error.message}`, 'error');
        }
        return 0;
    }
}

async function exportPaperDocs(files, useZip) {
    let exportedCount = 0;
    let zip = useZip ? new JSZip() : null;

    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');
    progressBar.max = files.length;
    progressBar.value = 0;

    for (let i = 0; i < files.length; i++) {
        if (!isExporting) {
            updateStatus('Export stopped by user.');
            break;
        }

        const entry = files[i];
        exportedCount += await exportSingleDoc(entry, zip);
        progressBar.value = i + 1;  // Update after each file

        // Update text with percentage
        const percent = Math.round(((i + 1) / files.length) * 100);
        progressText.innerHTML = `Progress: ${i + 1}/${files.length} files (${percent}%)`;

        await new Promise(resolve => setTimeout(resolve, 100));  // Rate limit delay
    }

    if (isExporting) {  // Only complete if not stopped
        if (useZip) {
            // Generate and download ZIP
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            const url = URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'dropbox-paper-exports.zip';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            updateStatus(`ZIP downloaded: ${exportedCount} files included!`);
        } else {
            updateStatus(`Exported ${exportedCount} Paper docs individually!`);
        }
    }

    // Reset progress and button after completion or stop
    progressBar.value = 0;
    progressText.innerHTML = '';
    resetButton();
}

function resetButton() {
    const btn = document.getElementById('exportBtn');
    btn.textContent = 'Start Export';
    btn.classList.remove('stop');
    btn.onclick = startExport;
    isExporting = false;
}

function stopExport() {
    isExporting = false;
}

async function startExport() {
    const token = document.getElementById('token').value.trim();
    if (!token) {
        updateStatus('Please enter your access token.', 'error');
        return;
    }

    const useZip = document.getElementById('useZip').checked;

    // Change button to stop
    const btn = document.getElementById('exportBtn');
    btn.textContent = 'Stop Export';
    btn.classList.add('stop');
    btn.onclick = stopExport;
    isExporting = true;

    try {
        dbx = new Dropbox.Dropbox({ accessToken: token });

        // Test auth
        await dbx.usersGetCurrentAccount();
        updateStatus(`Authenticated successfully. (ZIP: ${useZip})`);

        // Collect all files first for counting
        const files = await collectPaperFiles();
        if (files.length === 0) {
            updateStatus('No .paper files found in your Dropbox.', 'error');
            resetButton();
            return;
        }

        document.getElementById('progress-text').innerHTML = `Found ${files.length} files. Starting export...`;
        await exportPaperDocs(files, useZip);
    } catch (error) {
        console.error('Auth/Start error:', error);
        updateStatus(`Error: ${error.error_summary || error.message}. Check console (F12) for details.`, 'error');
        resetButton();
    }
}

function updateStatus(message, type = '') {
    const status = document.getElementById('status');
    status.innerHTML += `<p class="${type}">${new Date().toLocaleTimeString()}: ${message}</p>`;
    status.scrollTop = status.scrollHeight;
}
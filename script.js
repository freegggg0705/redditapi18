var gk_isXlsx = false;
var gk_xlsxFileLookup = {};
var gk_fileData = {};

function filledCell(cell) {
    return cell !== '' && cell != null;
}

function loadFileData(filename) {
    if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
        try {
            var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];

            var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
            var filteredData = jsonData.filter(row => row.some(filledCell));

            var headerRowIndex = filteredData.findIndex((row, index) =>
                row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
            );
            if (headerRowIndex === -1 || headerRowIndex > 25) {
                headerRowIndex = 0;
            }

            var subreddits = filteredData.slice(headerRowIndex + 1)
                .map(row => row[0])
                .filter(cell => cell && typeof cell === 'string' && cell.trim())
                .map(cell => cell.trim());
            return subreddits;
        } catch (e) {
            console.error(e);
            updateStatus('Error processing Excel file', true);
            return [];
        }
    }
    return [];
}

function updateStatus(message, isError = false) {
    const statusBar = document.getElementById('status-bar');
    statusBar.textContent = message;
    statusBar.style.background = isError ? '#dc3545' : '#007bff';
}

async function getAccessToken(clientId, clientSecret) {
    try {
        updateStatus('Fetching access token...');
        const response = await fetch('https://www.reddit.com/api/v1/access_token', {
            method: 'POST',
            headers: {
                'Authorization': 'Basic ' + btoa(`${clientId}:${clientSecret}`),
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: 'grant_type=client_credentials'
        });
        const data = await response.json();
        if (data.error) throw new Error(data.error);
        updateStatus('Access token retrieved');
        return data.access_token;
    } catch (error) {
        updateStatus(`Error getting access token: ${error.message}`, true);
        return null;
    }
}

async function fetchPostsBatch(clientId, clientSecret, subreddit, sort, limit, timeFilter, after = '') {
    try {
        const token = await getAccessToken(clientId, clientSecret);
        if (!token) return { mediaPosts: [], nonMediaPosts: [], after: '' };

        let mediaPosts = [];
        let nonMediaPosts = [];
        let baseUrl = `https://oauth.reddit.com/r/${subreddit}/${sort}.json?limit=${limit}`;
        if (sort === 'top' && timeFilter) {
            baseUrl += `&t=${timeFilter}`;
        }
        const url = after ? `${baseUrl}&after=${after}` : baseUrl;

        updateStatus(`Fetching ${limit} posts from r/${subreddit}...`);
        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        const data = await response.json();
        if (data.error) throw new Error(data.error);

        const posts = data.data.children.map(child => child.data);
        if (posts.length === 0) return { mediaPosts: [], nonMediaPosts: [], after: '' };

        posts.forEach(post => {
            if (post.url.match(/\.(gif|jpeg|jpg|png)$/i)) {
                mediaPosts.push(post);
            } else {
                nonMediaPosts.push(post);
            }
        });

        return { mediaPosts, nonMediaPosts, after: data.data.after };
    } catch (error) {
        updateStatus(`Error fetching posts from r/${subreddit}: ${error.message}`, true);
        return { mediaPosts: [], nonMediaPosts: [], after: '' };
    }
}

async function fetchPosts(clientId, clientSecret, subreddit, sort, limit, timeFilter) {
    let allMediaPosts = [];
    let allNonMediaPosts = [];
    let after = '';
    const maxAttempts = 10;
    let attempt = 0;

    while (allMediaPosts.length < limit && attempt < maxAttempts) {
        const { mediaPosts, nonMediaPosts, after: newAfter } = await fetchPostsBatch(
            clientId, clientSecret, subreddit, sort, limit, timeFilter, after
        );

        allMediaPosts.push(...mediaPosts);
        allNonMediaPosts.push(...nonMediaPosts);
        after = newAfter;
        attempt++;

        if (allMediaPosts.length > limit) {
            allNonMediaPosts.push(...allMediaPosts.slice(limit));
            allMediaPosts = allMediaPosts.slice(0, limit);
        }

        updateStatus(`Fetched ${allMediaPosts.length}/${limit} image posts from r/${subreddit}...`);

        if (allMediaPosts.length < limit && after) {
            const continueFetching = await new Promise(resolve => {
                const modal = document.getElementById('confirm-modal');
                const message = document.getElementById('confirm-message');
                const yesButton = document.getElementById('confirm-yes');
                const noButton = document.getElementById('confirm-no');

                message.textContent = `Only ${allMediaPosts.length} image posts found in r/${subreddit}. Fetch another ${limit} posts?`;
                modal.style.display = 'flex';

                yesButton.onclick = () => {
                    modal.style.display = 'none';
                    resolve(true);
                };
                noButton.onclick = () => {
                    modal.style.display = 'none';
                    resolve(false);
                };
            });

            if (!continueFetching) {
                updateStatus(`Stopped: Found ${allMediaPosts.length}/${limit} image posts in r/${subreddit}`, true);
                break;
            }
        } else if (!after) {
            updateStatus(`No more posts available in r/${subreddit}. Found ${allMediaPosts.length}/${limit} image posts`, true);
            break;
        }

        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (allMediaPosts.length >= limit) {
        updateStatus(`Successfully fetched ${allMediaPosts.length} image posts from r/${subreddit}`);
    }
    return { mediaPosts: allMediaPosts, nonMediaPosts: allNonMediaPosts };
}

async function displayMedia() {
    const clientId = document.getElementById('client-id').value.trim();
    const clientSecret = document.getElementById('client-secret').value.trim();
    let subredditInput = document.getElementById('subreddit-input').value.trim();
    const limitInput = parseInt(document.getElementById('limit-input').value) || 5;
    const sort = document.querySelector('.sort-button.active')?.dataset.sort || 'best';
    const timeFilter = sort === 'top' ? document.querySelector('.time-button.active')?.dataset.time || 'day' : null;

    if (!clientId || !clientSecret) {
        updateStatus('Please enter Client ID and Secret', true);
        return;
    }
    if (!subredditInput) {
        updateStatus('Please enter a subreddit or multireddit', true);
        return;
    }
    const limit = Math.min(Math.max(limitInput, 1), 100);

    const feedContainer = document.getElementById('feed-container');
    const nonMediaList = document.getElementById('non-media-items');
    const loadingSpinner = document.getElementById('loading-spinner');
    feedContainer.innerHTML = '';
    nonMediaList.innerHTML = '';
    loadingSpinner.style.display = 'block';

    const subreddits = subredditInput.split('+').map(s => s.trim()).filter(s => s);
    let allMediaPosts = [];
    let allNonMediaPosts = [];

    for (const subreddit of subreddits) {
        const { mediaPosts, nonMediaPosts } = await fetchPosts(clientId, clientSecret, subreddit, sort, limit, timeFilter);
        allMediaPosts.push(...mediaPosts);
        allNonMediaPosts.push(...nonMediaPosts);
    }

    if (allMediaPosts.length > limit) {
        allNonMediaPosts.push(...allMediaPosts.slice(limit));
        allMediaPosts = allMediaPosts.slice(0, limit);
    }

    loadingSpinner.style.display = 'none';

    allMediaPosts.forEach(post => {
        const feedItem = document.createElement('div');
        feedItem.className = 'feed-item';

        const img = document.createElement('img');
        img.className = 'thumbnail';
        img.src = post.url;
        img.alt = post.title;
        img.onerror = () => {
            img.className = 'thumbnail-placeholder';
            img.src = '';
            img.textContent = 'Image not available';
            const listItem = document.createElement('li');
            listItem.innerHTML = `Permalink: <a href="https://reddit.com${post.permalink}" target="_blank">${post.permalink}</a> | URL: <a href="${post.url}" target="_blank">${post.url}</a>`;
            nonMediaList.appendChild(listItem);
            feedItem.remove();
        };
        feedItem.appendChild(img);

        const title = document.createElement('a');
        title.className = 'title';
        title.href = `https://reddit.com${post.permalink}`;
        title.textContent = post.title.substring(0, 100);
        feedItem.appendChild(title);

        feedContainer.appendChild(feedItem);
    });

    allNonMediaPosts.forEach(post => {
        const listItem = document.createElement('li');
        listItem.innerHTML = `Permalink: <a href="https://reddit.com${post.permalink}" target="_blank">${post.permalink}</a> | URL: <a href="${post.url}" target="_blank">${post.url}</a>`;
        nonMediaList.appendChild(listItem);
    });

    if (allMediaPosts.length < limit) {
        updateStatus(`Only ${allMediaPosts.length} image posts found across subreddits`, true);
    } else {
        updateStatus(`Successfully fetched ${allMediaPosts.length} image posts`);
    }
}

function updateLayout() {
    const layout = document.querySelector('.layout-button.active')?.dataset.layout || 'grid';
    const columns = document.getElementById('columns-slider').value;
    const size = document.getElementById('size-slider').value;
    const feedContainer = document.getElementById('feed-container');

    feedContainer.className = layout;
    feedContainer.style.setProperty('--columns', columns);
    feedContainer.style.setProperty('--thumbnail-size', `${size}px`);
}

function setupEventListeners() {
    const timeFilterDiv = document.querySelector('.time-filter');

    document.querySelectorAll('.sort-button').forEach(button => {
        button.addEventListener('click', () => {
            document.querySelectorAll('.sort-button').forEach(btn => btn.classList.remove('active'));
            button.classList.add('active');
            timeFilterDiv.style.display = button.dataset.sort === 'top' ? 'flex' : 'none';
            if (button.dataset.sort === 'top') {
                document.querySelector('.time-button[data-time="day"]').classList.add('active');
            }
            displayMedia();
        });
    });

    document.querySelectorAll('.time-button').forEach(button => {
        button.addEventListener('click', () => {
            document.querySelectorAll('.time-button').forEach(btn => btn.classList.remove('active'));
            button.classList.add('active');
            displayMedia();
        });
    });

    document.querySelectorAll('.layout-button').forEach(button => {
        button.addEventListener('click', () => {
            document.querySelectorAll('.layout-button').forEach(btn => btn.classList.remove('active'));
            button.classList.add('active');
            updateLayout();
            displayMedia();
        });
    });

    document.getElementById('columns-slider').addEventListener('input', updateLayout);
    document.getElementById('size-slider').addEventListener('input', updateLayout);

    document.getElementById('client-id').addEventListener('change', displayMedia);
    document.getElementById('client-secret').addEventListener('change', displayMedia);
    document.getElementById('subreddit-input').addEventListener('change', displayMedia);
    document.getElementById('limit-input').addEventListener('change', displayMedia);

    document.getElementById('excel-file').addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;

        gk_isXlsx = true;
        const filename = file.name;
        gk_xlsxFileLookup[filename] = true;

        const reader = new FileReader();
        reader.onload = (e) => {
            const base64 = e.target.result.split(',')[1];
            gk_fileData[filename] = base64;
            const subreddits = loadFileData(filename);
            if (subreddits.length > 0) {
                document.getElementById('subreddit-input').value = subreddits.join('+');
                updateStatus(`Loaded ${subreddits.length} subreddits from Excel`);
                displayMedia();
            } else {
                updateStatus('No valid subreddits found in Excel file', true);
            }
        };
        reader.readAsDataURL(file);
    });
}

document.querySelector('.sort-button[data-sort="best"]').classList.add('active');
document.querySelector('.layout-button[data-layout="grid"]').classList.add('active');

setupEventListeners();
updateLayout();
updateStatus('Please enter Client ID and Secret', true);
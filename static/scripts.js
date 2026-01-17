document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const dropZone = document.getElementById('dropZone');
    const uploadText = document.getElementById('uploadText');
    const uploadSubtext = document.getElementById('uploadSubtext');
    const uploadIcon = document.getElementById('uploadIcon');
    const fileIcon = document.getElementById('fileIcon');
    const checkIcon = document.getElementById('checkIcon');
    const convertBtn = document.getElementById('convertBtn');
    const btnText = document.getElementById('btnText');
    const btnSpinner = document.getElementById('btnSpinner');
    const successMessage = document.getElementById('successMessage');
    const themeToggle = document.getElementById('themeToggle');
    const themeIcon = document.getElementById('themeIcon');

    // Ensure success message is hidden on page load
    successMessage.classList.add('hidden');

    // Theme Toggle
    const toggleTheme = () => {
        const isDark = document.documentElement.classList.toggle('dark');
        document.documentElement.classList.toggle('light', !isDark);
        localStorage.setItem('theme', isDark ? 'dark' : 'light');
        themeIcon.innerHTML = isDark ?
            `<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>` : // Moon icon
            `<path d="M12 3v2m0 14v2m9-9h-2m-14 0H3m16.95-6.95-1.42 1.42M6.34 17.66l-1.42 1.42M17.66 6.34l1.42-1.42M6.34 6.34l-1.42-1.42"/>`; // Sun icon
    };

    // Load saved theme
    const savedTheme = localStorage.getItem('theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    document.documentElement.classList.add(savedTheme);
    if (savedTheme === 'dark') {
        themeIcon.innerHTML = `<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>`;
    } else {
        themeIcon.innerHTML = `<path d="M12 3v2m0 14v2m9-9h-2m-14 0H3m16.95-6.95-1.42 1.42M6.34 17.66l-1.42 1.42M17.66 6.34l1.42-1.42M6.34 6.34l-1.42-1.42"/>`;
    }

    themeToggle.addEventListener('click', toggleTheme);

    // File Upload Handling
    function handleFileSelect(file) {
        if (!file) {
            resetUploadState();
            return;
        }
        if (!file.name.endsWith('.xlsx')) {
            uploadText.textContent = 'Invalid file type!';
            uploadSubtext.textContent = 'Please upload a .xlsx file';
            dropZone.classList.add('error');
            setTimeout(resetUploadState, 3000);
            return;
        }
        dropZone.classList.add('file-selected');
        dropZone.classList.remove('error');
        uploadIcon.classList.add('hidden');
        fileIcon.classList.remove('hidden');
        checkIcon.classList.remove('hidden');
        uploadText.textContent = file.name;
        uploadSubtext.textContent = 'File selected!';
        convertBtn.disabled = false;
        btnText.textContent = 'Convert & Download';
    }

    // Trigger file input click on drop zone click
    let isClicking = false;
    dropZone.addEventListener('click', (e) => {
        if (!isClicking) {
            isClicking = true;
            fileInput.click();
            setTimeout(() => { isClicking = false; }, 1000); // Prevent rapid re-triggering
        }
    });

    fileInput.addEventListener('change', (e) => {
        handleFileSelect(e.target.files[0]);
    });

    // Prevent re-triggering on cancel
    fileInput.addEventListener('click', (e) => {
        e.stopPropagation(); // Prevent dropZone click from re-triggering
    });

    // Drag and Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        fileInput.files = e.dataTransfer.files;
        handleFileSelect(file);
    });

    // Form Submission
    document.getElementById('uploadForm').addEventListener('submit', (e) => {
        btnText.textContent = 'Converting...';
        btnSpinner.classList.remove('hidden');
        convertBtn.disabled = true;
        successMessage.classList.add('hidden'); // Ensure hidden during submission
        setTimeout(() => {
            successMessage.classList.remove('hidden');
            setTimeout(() => {
                successMessage.classList.add('hidden');
                resetUploadState();
            }, 3000);
        }, 1000); // Adjust based on actual server response time
    });

    function resetUploadState() {
        uploadText.textContent = 'Click or drag Excel file here';
        uploadSubtext.textContent = 'Supports .xlsx files up to 10MB';
        uploadIcon.classList.remove('hidden');
        fileIcon.classList.add('hidden');
        checkIcon.classList.add('hidden');
        dropZone.classList.remove('file-selected', 'error', 'drag-over');
        convertBtn.disabled = true;
        btnText.textContent = 'Select File First';
        btnSpinner.classList.add('hidden');
        successMessage.classList.add('hidden');
        fileInput.value = '';
    }
});

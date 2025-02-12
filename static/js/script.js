let excelUploaded = false;
let pptUploaded = false;

function updateTransferButton() {
    const transferBtn = document.getElementById('transferBtn');
    transferBtn.disabled = !(excelUploaded && pptUploaded);
}

function showMessage(message, type) {
    const messageContainer = document.getElementById('messageContainer');
    messageContainer.innerHTML = `
        <div class="alert alert-${type} alert-dismissible fade show" role="alert">
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
        </div>
    `;
}

function showProgress(show) {
    const progressContainer = document.getElementById('progressContainer');
    progressContainer.classList.toggle('d-none', !show);
    if (show) {
        const progressBar = document.getElementById('progressBar');
        progressBar.style.width = '0%';
        let progress = 0;
        const interval = setInterval(() => {
            progress += 10;
            progressBar.style.width = `${progress}%`;
            if (progress >= 100) {
                clearInterval(interval);
            }
        }, 200);
    }
}

document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('excel_file', file);

    document.getElementById('excelFileName').textContent = file.name;
    document.getElementById('excelStatus').innerHTML = '<i class="fas fa-spinner fa-spin"></i> Uploading...';

    try {
        const response = await fetch('/upload-excel', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        
        if (response.ok) {
            document.getElementById('excelStatus').innerHTML = '<i class="fas fa-check text-success"></i> Uploaded';
            excelUploaded = true;
        } else {
            document.getElementById('excelStatus').innerHTML = '<i class="fas fa-times text-danger"></i> Failed';
            showMessage(data.error, 'danger');
        }
    } catch (error) {
        document.getElementById('excelStatus').innerHTML = '<i class="fas fa-times text-danger"></i> Failed';
        showMessage('Error uploading file', 'danger');
    }
    updateTransferButton();
});

document.getElementById('pptFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('ppt_file', file);

    document.getElementById('pptFileName').textContent = file.name;
    document.getElementById('pptStatus').innerHTML = '<i class="fas fa-spinner fa-spin"></i> Uploading...';

    try {
        const response = await fetch('/upload-powerpoint', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        
        if (response.ok) {
            document.getElementById('pptStatus').innerHTML = '<i class="fas fa-check text-success"></i> Uploaded';
            pptUploaded = true;
        } else {
            document.getElementById('pptStatus').innerHTML = '<i class="fas fa-times text-danger"></i> Failed';
            showMessage(data.error, 'danger');
        }
    } catch (error) {
        document.getElementById('pptStatus').innerHTML = '<i class="fas fa-times text-danger"></i> Failed';
        showMessage('Error uploading file', 'danger');
    }
    updateTransferButton();
});

document.getElementById('transferBtn').addEventListener('click', async () => {
    showProgress(true);
    try {
        const response = await fetch('/transfer', {
            method: 'POST'
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'modified_presentation.pptx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            showMessage('Transfer completed successfully! Downloading modified file...', 'success');
        } else {
            const data = await response.json();
            showMessage(data.error || 'Transfer failed', 'danger');
        }
    } catch (error) {
        showMessage('Error during transfer process', 'danger');
    }
    showProgress(false);
});

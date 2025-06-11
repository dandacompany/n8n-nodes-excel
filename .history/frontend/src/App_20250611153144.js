import React, { useState, useEffect } from 'react';
import axios from 'axios';
import './App.css';

const API_BASE_URL = 'http://localhost:3000';

// Modal Component
const Modal = ({ children, show, onClose }) => {
    if (!show) {
        return null;
    }
    return (
        <div className="modal-backdrop">
            <div className="modal-content">
                <button onClick={onClose} className="modal-close-btn">&times;</button>
                {children}
            </div>
        </div>
    );
};

function App() {
    const [files, setFiles] = useState([]);
    const [selectedFile, setSelectedFile] = useState('');
    const [fileData, setFileData] = useState(null);
    const [error, setError] = useState('');
    const [message, setMessage] = useState('');

    const [showCreateModal, setShowCreateModal] = useState(false);
    const [showModifyModal, setShowModifyModal] = useState(false);

    // Form states
    const [createFileName, setCreateFileName] = useState('');
    const [createCsvData, setCreateCsvData] = useState(`ID,Name,Value
1,ProductA,100
2,ProductB,150`);
    const [modifyCsvData, setModifyCsvData] = useState(`Status
Active`);
    const [modifySheet, setModifySheet] = useState('Sheet1');
    const [modifyOrigin, setModifyOrigin] = useState('A1');
    const [uploadFile, setUploadFile] = useState(null);

    const showMessage = (msg) => {
        setMessage(msg);
        setTimeout(() => setMessage(''), 3000);
    };

    const fetchFiles = async () => {
        try {
            const response = await axios.get(`${API_BASE_URL}/files`);
            setFiles(response.data);
            if (response.data.length > 0 && !response.data.includes(selectedFile)) {
                setSelectedFile(response.data[0]);
            } else if (response.data.length === 0) {
                setSelectedFile('');
                setFileData(null);
            }
        } catch (err) {
            setError('Could not fetch file list.');
        }
    };

    useEffect(() => {
        fetchFiles();
    }, []);

    const handleViewData = async () => {
        if (!selectedFile) return;
        try {
            const response = await axios.post(`${API_BASE_URL}/read`, { fileName: selectedFile });
            setFileData(response.data);
            setError('');
        } catch (err) {
            setError(`Error reading file: ${selectedFile}`);
            setFileData(null);
        }
    };

    const handleCreate = async (e) => {
        e.preventDefault();
        try {
            const payload = { fileName: createFileName, csvData: createCsvData };
            const response = await axios.post(`${API_BASE_URL}/create`, payload);
            showMessage(response.data.message);
            fetchFiles();
            setShowCreateModal(false);
        } catch (err) {
            setError(err.response ? err.response.data : 'Request failed.');
        }
    };

    const handleModify = async (e) => {
        e.preventDefault();
        try {
            const payload = { fileName: selectedFile, csvData: modifyCsvData, sheet: modifySheet, origin: modifyOrigin };
            const response = await axios.post(`${API_BASE_URL}/modify`, payload);
            showMessage(response.data.message);
            setShowModifyModal(false);
            handleViewData(); // Refresh data view
        } catch (err) {
            setError(err.response ? err.response.data : 'Request failed.');
        }
    };

    const handleDelete = async () => {
        if (!selectedFile) return;
        if (window.confirm(`Are you sure you want to delete ${selectedFile}?`)) {
            try {
                const response = await axios.post(`${API_BASE_URL}/delete`, { fileName: selectedFile });
                showMessage(response.data.message);
                fetchFiles();
            } catch (err) {
                setError(`Error deleting file: ${err.response ? err.response.data : err.message}`);
            }
        }
    };

    const handleUpload = async (e) => {
        e.preventDefault();
        if (!uploadFile) return;
        const formData = new FormData();
        formData.append('excel', uploadFile);
        try {
            const response = await axios.post(`${API_BASE_URL}/upload`, formData, {
                headers: { 'Content-Type': 'multipart/form-data' },
            });
            showMessage(response.data.message);
            fetchFiles();
        } catch (err) {
            setError(`Error uploading file: ${err.response ? err.response.data : err.message}`);
        }
    };

    const handleDownload = () => {
        if (!selectedFile) return;
        window.open(`${API_BASE_URL}/download/${selectedFile}`, '_blank');
    };

    const renderTable = () => {
        if (!fileData) return <p>Select a file and click "View Data" to see its content.</p>;
        if (fileData.length === 0) return <p>The file is empty.</p>;

        const originalHeaders = Object.keys(fileData[0] || {});
        const columnLetters = Array.from({ length: originalHeaders.length }, (_, i) => String.fromCharCode(65 + i));

        return (
            <table className="data-grid">
                <thead>
                    <tr>
                        <th className="index-header corner"></th>
                        {columnLetters.map(letter => (
                            <th key={letter} className="index-header">{letter}</th>
                        ))}
                    </tr>
                </thead>
                <tbody>
                    {/* Render original headers as the first data row */}
                    <tr>
                        <td className="index-header">1</td>
                        {originalHeaders.map(header => (
                            <td key={header} className="original-header-cell">{header}</td>
                        ))}
                    </tr>
                    {/* Render the actual data rows */}
                    {fileData.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                            <td className="index-header">{rowIndex + 2}</td>
                            {originalHeaders.map(header => (
                                <td key={header}>{String(row[header])}</td>
                            ))}
                        </tr>
                    ))}
                </tbody>
            </table>
        );
    };

    return (
        <div className="App">
            <header className="App-header"><h1>Excel-as-DB Interface</h1></header>
            {error && <p className="error-message" onClick={() => setError('')}>{error}</p>}
            {message && <p className="success-message" onClick={() => setMessage('')}>{message}</p>}
            <main>
                <section className="file-manager">
                    <h2>File Management</h2>
                    <div className="file-controls">
                        <select value={selectedFile} onChange={e => setSelectedFile(e.target.value)}>
                            {files.map(f => <option key={f} value={f}>{f}</option>)}
                        </select>
                        <button onClick={handleViewData} disabled={!selectedFile}>View Data</button>
                        <button onClick={() => setShowCreateModal(true)} className="create-btn">Create New</button>
                    </div>
                    <div className="file-actions">
                        <button onClick={handleDownload} disabled={!selectedFile}>Download</button>
                        <button onClick={() => setShowModifyModal(true)} disabled={!selectedFile}>Modify</button>
                        <button onClick={handleDelete} disabled={!selectedFile} className="delete-btn">Delete</button>
                    </div>
                    <form className="upload-form" onSubmit={handleUpload}>
                        <input type="file" onChange={e => setUploadFile(e.target.files[0])} accept=".xlsx" />
                        <button type="submit" disabled={!uploadFile}>Upload File</button>
                    </form>
                </section>
                <section className="data-viewer">
                    <h2>Data Viewer: {selectedFile || 'None'}</h2>
                    <div className="table-container">{renderTable()}</div>
                </section>
            </main>

            <Modal show={showCreateModal} onClose={() => setShowCreateModal(false)}>
                <form onSubmit={handleCreate}>
                    <h2>Create New Excel File</h2>
                    <label>File Name:</label>
                    <input type="text" value={createFileName} onChange={e => setCreateFileName(e.target.value)} placeholder="e.g., my_data.xlsx" required />
                    <label>Data (CSV Format):</label>
                    <textarea value={createCsvData} onChange={e => setCreateCsvData(e.target.value)} rows="5" required />
                    <button type="submit">Create File</button>
                </form>
            </Modal>

            <Modal show={showModifyModal} onClose={() => setShowModifyModal(false)}>
                <form onSubmit={handleModify}>
                    <h2>Modify: {selectedFile}</h2>
                    <label>Data (CSV Format):</label>
                    <textarea value={modifyCsvData} onChange={e => setModifyCsvData(e.target.value)} rows="5" required />
                    <label>Sheet Name (optional):</label>
                    <input type="text" value={modifySheet} onChange={e => setModifySheet(e.target.value)} />
                    <label>Origin (e.g., A1, C5):</label>
                    <input type="text" value={modifyOrigin} onChange={e => setModifyOrigin(e.target.value)} />
                    <button type="submit">Apply Modifications</button>
                </form>
            </Modal>
        </div>
    );
}

export default App;

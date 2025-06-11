import React, { useState } from 'react';
import axios from 'axios';
import './App.css';

const API_BASE_URL = 'http://localhost:3000';

function App() {
    // Create state
    const [createFileName, setCreateFileName] = useState('new_data.xlsx');
    const [createData, setCreateData] = useState('[{"Name": "Alice", "ID": 1}, {"Name": "Bob", "ID": 2}]');
    const [createHeaders, setCreateHeaders] = useState('ID,Name');

    // Read state
    const [readFileName, setReadFileName] = useState('new_data.xlsx');

    // Modify state
    const [modifyFileName, setModifyFileName] = useState('new_data.xlsx');
    const [modifyData, setModifyData] = useState('[["Status"], ["Active"]]');
    const [modifySheet, setModifySheet] = useState('Sheet1');
    const [modifyOrigin, setModifyOrigin] = useState('C1');

    // API Response state
    const [apiResponse, setApiResponse] = useState(null);

    const handleApiCall = async (endpoint, payload) => {
        try {
            const response = await axios.post(`${API_BASE_URL}${endpoint}`, payload);
            setApiResponse({ status: 'Success', data: response.data });
        } catch (error) {
            setApiResponse({ status: 'Error', data: error.response ? error.response.data : error.message });
        }
    };

    const handleCreate = (e) => {
        e.preventDefault();
        let data;
        try {
            data = JSON.parse(createData);
        } catch (error) {
            setApiResponse({ status: 'Error', data: 'Invalid JSON format for data.' });
            return;
        }
        const headers = createHeaders.split(',').map(h => h.trim()).filter(Boolean);
        handleApiCall('/create', { fileName: createFileName, data, headers: headers.length ? headers : undefined });
    };

    const handleRead = (e) => {
        e.preventDefault();
        handleApiCall('/read', { fileName: readFileName });
    };

    const handleModify = (e) => {
        e.preventDefault();
        let data;
        try {
            data = JSON.parse(modifyData);
        } catch (error) {
            setApiResponse({ status: 'Error', data: 'Invalid JSON format for data.' });
            return;
        }
        handleApiCall('/modify', { fileName: modifyFileName, data, sheet: modifySheet, origin: modifyOrigin });
    };

    return (
        <div className="App">
            <header className="App-header">
                <h1>Excel DB API Tester</h1>
            </header>
            <main className="App-main">
                <div className="api-section">
                    <h2>Create Excel File</h2>
                    <form onSubmit={handleCreate}>
                        <label>File Name:</label>
                        <input type="text" value={createFileName} onChange={(e) => setCreateFileName(e.target.value)} />
                        <label>Data (JSON Array of Objects):</label>
                        <textarea value={createData} onChange={(e) => setCreateData(e.target.value)} rows="5" />
                        <label>Headers (comma-separated, optional):</label>
                        <input type="text" value={createHeaders} onChange={(e) => setCreateHeaders(e.target.value)} />
                        <button type="submit">Create File</button>
                    </form>
                </div>

                <div className="api-section">
                    <h2>Read Excel File</h2>
                    <form onSubmit={handleRead}>
                        <label>File Name:</label>
                        <input type="text" value={readFileName} onChange={(e) => setReadFileName(e.target.value)} />
                        <button type="submit">Read File</button>
                    </form>
                </div>

                <div className="api-section">
                    <h2>Modify Excel File</h2>
                    <form onSubmit={handleModify}>
                        <label>File Name:</label>
                        <input type="text" value={modifyFileName} onChange={(e) => setModifyFileName(e.target.value)} />
                        <label>Data (JSON 2D Array):</label>
                        <textarea value={modifyData} onChange={(e) => setModifyData(e.target.value)} rows="5" />
                        <label>Sheet Name (optional):</label>
                        <input type="text" value={modifySheet} onChange={(e) => setModifySheet(e.target.value)} />
                        <label>Origin (e.g., A1, C5):</label>
                        <input type="text" value={modifyOrigin} onChange={(e) => setModifyOrigin(e.target.value)} />
                        <button type="submit">Modify File</button>
                    </form>
                </div>
            </main>
            <footer className="App-footer">
                <h2>API Response</h2>
                {apiResponse && (
                    <div className={`response ${apiResponse.status.toLowerCase()}`}>
                        <strong>{apiResponse.status}</strong>
                        <pre>{JSON.stringify(apiResponse.data, null, 2)}</pre>
                    </div>
                )}
            </footer>
        </div>
    );
}

export default App; 
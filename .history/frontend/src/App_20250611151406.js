import React, { useState, useEffect } from 'react';
import axios from 'axios';
import './App.css';

const API_BASE_URL = 'http://localhost:3000';

function App() {
    const [files, setFiles] = useState([]);
    const [selectedFile, setSelectedFile] = useState('');
    const [fileData, setFileData] = useState(null);
    const [error, setError] = useState('');

    // Fetch file list on component mount
    const fetchFiles = async () => {
        setError('');
        try {
            const response = await axios.get(`${API_BASE_URL}/files`);
            setFiles(response.data);
            if (response.data.length > 0) {
                // If there is no selected file or the selected file is not in the new list, reset it.
                if (!selectedFile || !response.data.includes(selectedFile)) {
                    setSelectedFile(response.data[0]);
                }
            } else {
                setSelectedFile('');
            }
        } catch (err) {
            setError('서버에서 파일 목록을 가져오는데 실패했습니다.');
            console.error(err);
        }
    };

    useEffect(() => {
        fetchFiles();
    }, []);

    const handleFileSelect = (e) => {
        setSelectedFile(e.target.value);
        setFileData(null); // Clear previous data when new file is selected
    };

    const handleViewData = async () => {
        if (!selectedFile) {
            setError('데이터를 조회할 파일을 선택하세요.');
            return;
        }
        setError('');
        setFileData(null);
        try {
            const response = await axios.post(`${API_BASE_URL}/read`, { fileName: selectedFile });
            setFileData(response.data);
        } catch (err) {
            setError(`'${selectedFile}' 파일 데이터를 읽는 중 오류가 발생했습니다.`);
            console.error(err);
        }
    };

    const renderTable = () => {
        if (!fileData) {
            return <p>표시할 데이터가 없습니다. 파일을 선택하고 'View Data'를 클릭하세요.</p>;
        }
        if (fileData.length === 0) {
            return <p>파일에 데이터가 없습니다.</p>;
        }

        const headers = Object.keys(fileData[0]);
        return (
            <table>
                <thead>
                    <tr>
                        {headers.map(header => <th key={header}>{header}</th>)}
                    </tr>
                </thead>
                <tbody>
                    {fileData.map((row, index) => (
                        <tr key={index}>
                            {headers.map(header => <td key={header}>{String(row[header])}</td>)}
                        </tr>
                    ))}
                </tbody>
            </table>
        );
    };

    return (
        <div className="App">
            <header className="App-header">
                <h1>Excel-as-DB Interface</h1>
            </header>
            <main>
                <section className="file-manager">
                    <h2>File Explorer</h2>
                    <div className="file-controls">
                        {files.length > 0 ? (
                            <>
                                <select value={selectedFile} onChange={handleFileSelect}>
                                    {files.map(file => (
                                        <option key={file} value={file}>{file}</option>
                                    ))}
                                </select>
                                <button onClick={handleViewData}>View Data</button>
                            </>
                        ) : (
                            <p>data 폴더에 조회할 파일이 없습니다.</p>
                        )}
                        <button onClick={fetchFiles} className="refresh-btn">Refresh List</button>
                    </div>
                </section>
                <section className="data-viewer">
                    <h2>Data Viewer: {selectedFile || 'None'}</h2>
                    {error && <p className="error-message">{error}</p>}
                    <div className="table-container">
                        {renderTable()}
                    </div>
                </section>
            </main>
        </div>
    );
}

export default App;

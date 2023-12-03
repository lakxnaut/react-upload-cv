import React, { useCallback, useState } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import FileSaver from 'file-saver';

const App = () => {
  const [jsonData, setJsonData] = useState([]);

  const onDrop = useCallback((acceptedFiles) => {
    const file = acceptedFiles[0];
    const reader = new FileReader();

    reader.onload = async (event) => {
      try {
        const data = event.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 'firstRow' });

        // Log the JSON data to the console
        console.log('JSON Data:', jsonData);

        // Set the JSON data to the state
        setJsonData(jsonData);
      } catch (error) {
        console.error('Error reading Excel file:', error);
      }
    };

    reader.readAsBinaryString(file);
  }, []);

  const downloadDocx = (candidate) => {
    try {
      // Dynamic template with placeholders
      const template = `
  Candidate Name: ${candidate.name}
  Personal Information:
    Gender: ${candidate.Gen}
    Nationality: ${candidate.Nationality}
    Date of Birth: ${candidate.Birthday}
    Phone Number: ${candidate.Tel}
    Email Address: ${candidate.Email}
  Driving License: ${candidate.License}
  Disponibility: ${candidate.Disponibility}
  Mobility: ${candidate.Mobility}
  Image : ${candidate.Image}
   `;
  
      // Create a new PizZip instance
      const zip = new PizZip();
      zip.load(template);
  
      const doc = new Docxtemplater();
      doc.loadZip(zip);
  
      // Set the data for docxtemplater
      doc.setData(candidate);
  
      // Check for empty or invalid data
      if (Object.keys(candidate).length === 0) {
        console.error('Invalid data for rendering the document');
        return;
      }
  
      try {
        // Render the document
        doc.render();
      } catch (renderError) {
        console.error('Error during rendering:', renderError);
        return;
      }
  
      // Get the generated document
      const generatedDoc = doc.getZip().generate({ type: 'blob' });
  
      // Save the document
      FileSaver.saveAs(generatedDoc, `document_${candidate['Candidate Name']}.docx`);
    } catch (error) {
      console.error('Error generating Word document:', error);
    }
  };
  

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: '.xls, .xlsx',
  });

  return (
    <div style={{ textAlign: 'center', marginTop: '50px' }}>
      <h2>React File Upload, Convert to JSON, and Download DOCX</h2>
      <div {...getRootProps()} style={dropzoneStyles}>
        <input {...getInputProps()} />
        {isDragActive ? (
          <p>Drop the files here ...</p>
        ) : (
          <p>Drag 'n' drop an Excel file here, or click to select a file</p>
        )}
      </div>
      {jsonData.length > 0 && (
        <div style={{ marginTop: '20px' }}>
          <h3>Converted JSON Data:</h3>
          <ul>
            {jsonData.map((object, index) => (
              <li key={index}>
                {object.Name} -{' '}
                <button onClick={() => downloadDocx(object)}>Download DOCX</button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
};

const dropzoneStyles = {
  width: '200px',
  height: '150px',
  border: '2px dashed #cccccc',
  borderRadius: '4px',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  cursor: 'pointer',
  margin: 'auto',
};

export default App;

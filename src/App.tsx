import { useRef, useEffect, useState } from 'react';
import {
  DocumentEditorContainerComponent
} from '@syncfusion/ej2-react-documenteditor';
import { ButtonComponent } from '@syncfusion/ej2-react-buttons';
import './App.css';

function App() {
  const editorRef1 = useRef<DocumentEditorContainerComponent | null>(null);
  const editorRef2 = useRef<DocumentEditorContainerComponent | null>(null);

  const [originalFile, setOriginalFile] = useState<File | null>(null);
  const [revisedFile, setRevisedFile] = useState<File | null>(null);
  const [showResult, setShowResult] = useState<boolean>(true);
  const [compareClicked, setCompareClicked] = useState(false);
  // Store whether editors have document open or are empty
  const [editorsLoaded, setEditorsLoaded] = useState(false);

  //Utilise the correct Service URL where the Import and CompareDocument Web APIs are hosted.
  //If you choose to use localhost, please update the URL below to 'http://localhost:62870/api/DocumentEditor/'; 
  let serviceUrl: string = 'https://document.syncfusion.com/web-services/docx-editor/api/documenteditor/'; 

  // File input refs for hidden file input
  const fileInputRef1 = useRef<HTMLInputElement>(null);
  const fileInputRef2 = useRef<HTMLInputElement>(null);

  const openFileInEditor = async (
    file: File,
    editorRef: React.RefObject<DocumentEditorContainerComponent | null>
  ) => {
    if (file.name.endsWith('.sfdt')) {
      const reader = new FileReader();
      reader.onload = e => {
        const docData = e.target?.result as string | undefined;
        if (docData && editorRef.current?.documentEditor) {
          editorRef.current.documentEditor.open(docData);
        }
      };
      reader.readAsText(file);
    }
    else {
      const formData = new FormData();
      formData.append('file', file);
      try {
        const response = await fetch(serviceUrl + 'Import', {
          method: 'POST',
          body: formData
        });
        const sfdtString = await response.text();
        let sfdtObject: any = null;
        try {
          sfdtObject = JSON.parse(sfdtString);
        } catch (e) {
          alert("Failed to parse SFDT JSON string from backend.");
        }
        if (sfdtObject && editorRef.current?.documentEditor) {
          editorRef.current.documentEditor.open(sfdtObject);
        } else {
          alert("SFDT object missing or editor not ready.");
        }
      } catch (e) {
        alert('Failed to open .docx file. Make sure your backend supports DOCX import. See server docs.');
      }
    }
  };

  const onCompare = async () => {
    setCompareClicked(false); // Reset to re-mount editors
    setEditorsLoaded(false);  // Reset editors' loaded state
    setTimeout(async () => {
      setCompareClicked(true);
      setEditorsLoaded(true); // Editors will now load documents
      if (showResult) {
        if (originalFile) await openFileInEditor(originalFile, editorRef1);
        if (originalFile && revisedFile) {
          await loadComparedDocumentAndOpen(editorRef2 as React.RefObject<DocumentEditorContainerComponent>, originalFile, revisedFile);
        }
      } else {
        if (originalFile) await openFileInEditor(originalFile, editorRef1);
        if (revisedFile) await openFileInEditor(revisedFile, editorRef2);
      }
    }, 0);
  };

  // open the compared document in the editor (right editor)
  async function loadComparedDocumentAndOpen(
    editorRef2: React.RefObject<DocumentEditorContainerComponent>,
    originalFile: File,
    revisedFile: File
  ) {
    const formData = new FormData();
    formData.append("originalFile", originalFile);
    formData.append("revisedFile", revisedFile);
    formData.append("author", "Author");
    formData.append("dateTime", new Date().toISOString());
    const response = await fetch(serviceUrl + 'CompareDocuments', {
      method: 'POST',
      body: formData
    });
    if (!response.ok) {
      alert('Failed to fetch compared document.');
      return;
    }
    const sfdtString = await response.text();
    let sfdtObject: any = null;
    try {
      sfdtObject = JSON.parse(sfdtString);
    } catch {
      alert("Compare API did not return valid JSON.");
      return;
    }
    if (sfdtObject && editorRef2.current?.documentEditor) {
      editorRef2.current.documentEditor.open(sfdtObject);
    } else {
      alert("Compare API did not return SFDT document or editor2 not ready");
    }
  }

  // Download the revised/result file (right editor)
  const onDownload = () => {
    if (editorRef2.current?.documentEditor) {
      editorRef2.current.documentEditor.save('Result', 'Docx');
    }
  };

  useEffect(() => {
    const editor1 = editorRef1.current?.documentEditor;
    const editor2 = editorRef2.current?.documentEditor;
    if (editor1 && editor2) {
      debugger;
      editor1.fitPage('FitPageWidth');
      editor2.fitPage('FitPageWidth')
      editor1.viewChange = () => {
        const pos = editor1.selection.getScrollPosition();
        editor2.selection.setScrollPosition(pos);
      };
      editor2.viewChange = () => {
        const pos = editor2.selection.getScrollPosition();
        editor1.selection.setScrollPosition(pos);
      };
    }
  });

  return (
    <div className="App" style={{ width: '100%' }}>

      {/* Shared Container */}
      <div
        style={{
          width: 760,
          margin: '20px auto 10px auto'
        }}
      >
        {/* Upload Row */}
        <div style={{ display: 'flex', gap: 6 }}>

          {/* Original */}
          <div
            style={{
              width: 360,
              background: '#FAFAFA',
              borderRadius: 6,
              padding: 14,
              border: '1px dashed #ccc'
            }}
          >
            {/* Top row */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <span>Original document</span>

              <label
                style={{
                  padding: '4px 10px',
                  border: '1px solid #ccc',
                  borderRadius: 4,
                  background: '#fff',
                  cursor: 'pointer'
                }}
              >
                Choose file
                <input
                  type="file"
                  style={{ display: 'none' }}
                  accept=".docx,.sfdt"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                      setOriginalFile(file);
                      setCompareClicked(false);
                      setEditorsLoaded(false);
                    }
                  }}
                />
              </label>
            </div>

            {/* Filename + Delete */}
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginTop: 6,
                fontSize: 12,
                color: '#666'
              }}
            >
              <span>
                {originalFile
                  ? `${originalFile.name} (${Math.round(originalFile.size / 1024)} KB)`
                  : 'File supports sfdt, docx'}
              </span>

              {originalFile && (
                <span
                  className="e-icons e-trash"
                  onClick={() => setOriginalFile(null)}
                  style={{
                    cursor: 'pointer'
                  }}
                  title="Remove file"
                />
              )}
            </div>
          </div>

          {/* Revised */}
          <div
            style={{
              width: 360,
              background: '#FAFAFA',
              borderRadius: 6,
              padding: 14,
              border: '1px dashed #ccc'
            }}
          >
            {/* Top row */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <span>Revised document</span>

              <label
                style={{
                  padding: '4px 10px',
                  border: '1px solid #ccc',
                  borderRadius: 4,
                  background: '#fff',
                  cursor: 'pointer'
                }}
              >
                Choose file
                <input
                  type="file"
                  style={{ display: 'none' }}
                  accept=".docx,.sfdt"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                      setRevisedFile(file);
                      setCompareClicked(false);
                      setEditorsLoaded(false);
                    }
                  }}
                />
              </label>
            </div>

            {/* Filename + Delete */}
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginTop: 6,
                fontSize: 12,
                color: '#666'
              }}
            >
              <span>
                {revisedFile
                  ? `${revisedFile.name} (${Math.round(revisedFile.size / 1024)} KB)`
                  : 'File supports sfdt, docx'}
              </span>

              {revisedFile && (
                <span
                  className="e-icons e-trash"
                  onClick={() => setRevisedFile(null)}
                  style={{
                    cursor: 'pointer'
                  }}
                  title="Remove file"
                />
              )}
            </div>
          </div>
        </div>

        {/* Controls */}
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginTop: 12
          }}
        >
          {/* Checkbox */}
          <div>
            <input
              type="checkbox"
              checked={showResult}
              onChange={(e) => setShowResult(e.target.checked)}
              style={{ marginRight: 6 }}
            />
            <span>Show comparison results with tracked changes</span>
          </div>

          {/* Buttons */}
          <div style={{ display: 'flex', gap: 10 }}>
            <ButtonComponent
              cssClass="e-primary"
              disabled={!originalFile || !revisedFile}
              onClick={onCompare}
            >
              Compare
            </ButtonComponent>

            <ButtonComponent
              cssClass="e-outline e-flat e-primary"
              disabled={!compareClicked}
              onClick={onDownload}
              iconCss="e-icons e-download"
            >
              Results
            </ButtonComponent>
          </div>
        </div>
      </div>

      {/* Editors */}
      <div
        style={{
          display: 'flex',
          gap: 6,
          height: '600px',
          padding: '10px'
        }}
      >
        {/* Left Editor */}
        <div
          style={{
            width: '50%',
            display: 'flex',
            flexDirection: 'column',
            border: '1px solid #ddd',
            overflow: 'hidden'
          }}
        >
          <div className="e-de-ctn-title" style={{ padding: '8px', fontWeight: 600 }}>
            Original Document
          </div>

          <div style={{ flex: 1 }}>
            <DocumentEditorContainerComponent
              id="editor1"
              ref={editorRef1}
              serviceUrl={serviceUrl}
              height="100%"
              width="100%"
              enableToolbar={false}
              showPropertiesPane={false}
            />
          </div>
        </div>

        {/* Right Editor */}
        <div
          style={{
            width: '50%',
            display: 'flex',
            flexDirection: 'column',
            border: '1px solid #ddd',
            overflow: 'hidden'
          }}
        >
          <div className="e-de-ctn-title" style={{ padding: '8px', fontWeight: 600 }}>
            {showResult
              ? 'Result Document (with tracked changes)'
              : 'Revised Document'}
          </div>

          <div style={{ flex: 1 }}>
            <DocumentEditorContainerComponent
              id="editor2"
              ref={editorRef2}
              serviceUrl={serviceUrl}
              height="100%"
              width="100%"
              enableToolbar={false}
              showPropertiesPane={false}
            />
          </div>
        </div>
      </div>

    </div>
  );
}

export default App;
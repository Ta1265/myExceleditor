import React, { useState } from 'react';
import List from './List';
import FileUploader from './FileUploader';
import UploadFilesService from '../uploadService'
import fileDownload from 'js-file-download'
import axios from 'axios'


function App(): JSX.Element {
  const [state, changeState] = useState<string>('Hello World');
  const [arFiles, addArFiles] = useState<Array<File>>([]);
  const [ledgerFiles, addLedgerFiles] = useState<Array<File>>([]);

  const onFormatFilesClick = () => {
    axios.get('/format', {responseType: 'blob'})
      .then(({data}) => {
        const downloadUrl = window.URL.createObjectURL(new Blob([data]));
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.setAttribute('download', 'file.xlsx'); //any other extension
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.location.reload();
      })
      .catch((err) => console.log(err.messege));
  }
  const onClickReset = () => {
    axios.get('/resetfiles')
      .then((res) => {
        window.location.reload();
      })
  }

  return (
    <div>
      <FileUploader 
        typeTitle="Add AR Files here"
        filetype="ar"
      />
      <FileUploader 
        typeTitle="Trust Ledger Files here"
        filetype="ledger"
      />
      <br>
      </br>
      <button onClick={onFormatFilesClick}>
        Format Documents
      </button>
      <button onClick={onClickReset}>
        Reset 
      </button>
    </div>
  );
}

export default App;

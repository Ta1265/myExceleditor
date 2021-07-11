import axios from 'axios';

class UploadFilesService {
  upload(file:File, type:string, onUploadProgress:any) {
    let formData = new FormData();

    formData.append("file", file);

    return axios.post(`/upload${type}`, formData, {
      headers: {
        "Content-Type": "application/vnd.ms-excel",
      },
      onUploadProgress,
    });
  }

  getFiles() {
    return axios.get("/files");
  }
}

export default new UploadFilesService();
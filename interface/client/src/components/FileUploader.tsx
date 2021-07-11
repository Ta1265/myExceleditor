import { createPublicKey } from 'crypto';
import React from 'react';
import Presentation from './Presentation';
import UploadFilesService from '../uploadService'

type State = {
  dragging: boolean;
  files: Array<File> | null;
}

type Props = {
  typeTitle: string;
  filetype: string;
}


export default class FileUploader extends React.Component<Props, State> {
  static counter = 0;
  FileUploaderInput: React.RefObject<HTMLInputElement>;

  constructor(props: Props) {
    super(props);
    this.state = { dragging: false, files: [] };
    this.FileUploaderInput = React.createRef();
  }

  async uploadFile(file: File) {
    const { filetype } = this.props
    return await UploadFilesService.upload(file, filetype, () => {})
      .then((response: any) => {
        console.log("success", response);
        const morefiles = this.state.files;
        morefiles?.push(file);
        this.setState({files : morefiles})
      })
      .catch((error) => {
        console.log("error uploading files", error.message);
        alert(`error uploading files ${error.message}`);
      })
  }



  dragEventCounter = 0;
  dragenterListener = (event: React.DragEvent<HTMLDivElement>) => {
    this.overrideEventDefaults(event);
    this.dragEventCounter++;
    if (event.dataTransfer.items && event.dataTransfer.items[0]) {
      this.setState({ dragging: true });
    } else if (
      event.dataTransfer.types &&
      event.dataTransfer.types[0] === "Files"
    ) {
      // This block handles support for IE - if you're not worried about
      // that, you can omit this
      this.setState({ dragging: true });
    }
  };

  dragleaveListener = (event: React.DragEvent<HTMLDivElement>) => {
    this.overrideEventDefaults(event);
    this.dragEventCounter--;

    if (this.dragEventCounter === 0) {
      this.setState({ dragging: false });
    }
  };

  dropListener = (event: React.DragEvent<HTMLDivElement>) => {
    this.overrideEventDefaults(event);
    this.dragEventCounter = 0;
    this.setState({ dragging: false });

    if (event.dataTransfer.files && event.dataTransfer.files[0]) {
      for(let i = 0; i < event.dataTransfer.files.length; i += 1) {
        this.uploadFile(event.dataTransfer.files[i]);
      }
    }
  };

  overrideEventDefaults = (event: Event | React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
  };

  onSelectFileClick = () => {
  };

  onFileChanged = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files[0]) {
      for(let i = 0; i < event.target.files.length; i +=1 ) {
        this.uploadFile(event.target.files[i]);
      }
    }
  };

  componentDidMount() {
    window.addEventListener("dragover", (event: Event) => {
      this.overrideEventDefaults(event);
    });
    window.addEventListener("drop", (event: Event) => {
      this.overrideEventDefaults(event);
    });
  }

  componentWillUnmount() {
    window.removeEventListener("dragover", this.overrideEventDefaults);
    window.removeEventListener("drop", this.overrideEventDefaults);
  }

  render() {
    const { typeTitle } = this.props;
    return (
      <Presentation
        typeTitle={typeTitle}
        dragging={this.state.dragging}
        files={this.state.files}
        onSelectFileClick={this.onSelectFileClick}
        onDrag={this.overrideEventDefaults}
        onDragStart={this.overrideEventDefaults}
        onDragEnd={this.overrideEventDefaults}
        onDragOver={this.overrideEventDefaults}
        onDragEnter={this.dragenterListener}
        onDragLeave={this.dragleaveListener}
        onDrop={this.dropListener}
      >
        <input
          id="fileElem"
          ref={this.FileUploaderInput}
          type="file"
          className="file-uploader__input"
          onChange={this.onFileChanged}
          accept="application/vnd.ms-excel"
          multiple
        />
      </Presentation>
    );
  }
}
import React from 'react';

type PresentationalProps = {
  typeTitle: string;
  dragging: boolean;
  files: Array<File> | null;
  onSelectFileClick: () => void;
  onDrag: (event: React.DragEvent<HTMLDivElement>) => void;
  onDragStart: (event: React.DragEvent<HTMLDivElement>) => void;
  onDragEnd: (event: React.DragEvent<HTMLDivElement>) => void;
  onDragOver: (event: React.DragEvent<HTMLDivElement>) => void;
  onDragEnter: (event: React.DragEvent<HTMLDivElement>) => void;
  onDragLeave: (event: React.DragEvent<HTMLDivElement>) => void;
  onDrop: (event: React.DragEvent<HTMLDivElement>) => void;
};

const Presentation: React.FunctionComponent<PresentationalProps> = props => {
  const {
    typeTitle,
    dragging,
    files,
    onSelectFileClick,
    onDrag,
    onDragStart,
    onDragEnd,
    onDragOver,
    onDragEnter,
    onDragLeave,
    onDrop
  } = props;

  let uploaderClasses = "file-uploader";
  if (dragging) {
    uploaderClasses += " file-uploader--dragging";
  }

  const fileNames = files ? files.map((file) => file.name) : ["No Files Uploaded!"];

  return (
    <div
      className={uploaderClasses}
      onDrag={onDrag}
      onDragStart={onDragStart}
      onDragEnd={onDragEnd}
      onDragOver={onDragOver}
      onDragEnter={onDragEnter}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <div className="file-uploader__contents">
        <span className="typeTitle">{typeTitle}</span>
        <span className="file-uploader__file-name">
          <ul>
            {fileNames.map((name, index) => <li key={index}>{name}</li>)}
          </ul>
        </span>
        <span>Drag & Drop File</span>
        <span>or</span>
        <span onClick={onSelectFileClick}>
          {props.children}
        </span>
      </div>
    </div>
  );
};

export default Presentation;
import React from 'react';
import axios from 'axios';

interface Props {
  nothing: string;
}

interface State {
  listItems: Array<string>;
}

export default class List extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount(): void {
    console.log('hi');
    axios.get<Array<string>>('/listitems')
      .then((res) => {
        console.log('what did i get?', res.data);
        this.setState({
          listItems: res.data,
        });
      })
      .catch((error: Error) => {
        console.error(error.message);
      });
  }

  render(): JSX.Element {
    const { listItems } = this.state;
    return (
      <div>
        <ul>{listItems.map((item) => (<li>{item}</li>))}</ul>
      </div>
    );
  }
}

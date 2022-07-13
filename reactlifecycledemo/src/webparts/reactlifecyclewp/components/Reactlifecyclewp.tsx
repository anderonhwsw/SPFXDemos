import * as React from 'react';
import styles from './Reactlifecyclewp.module.scss';
import { IReactlifecyclewpProps } from './IReactlifecyclewpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IReactlifecyclewpWebPartProps } from '../ReactlifecyclewpWebPart';

export interface IReactlifecyclewpState {
  stageTitle: string;
}

export default class Reactlifecyclewp extends React.Component<IReactlifecyclewpProps, IReactlifecyclewpState> {

  public constructor (props: IReactlifecyclewpProps, state: IReactlifecyclewpState){
    super(props);
    this.state = {
      stageTitle: 'component Constructor has been called'
    };

    this.updateState = this.updateState.bind(this);

    console.log('Stage Title from Constructor: ' + this.state.stageTitle);
  }

  public componentWillMount() {
    console.log('component will mount has been called');
  }

  public componentDidMount() {
    console.log('Stage Title from componentDidMount : ' + this.state.stageTitle);
    this.setState({
      stageTitle: 'componentDidMount has been called'
    });
  }

  public updateState(){
    this.setState({
      stageTitle: 'changeState has been called'
    })
  }

  public render(): React.ReactElement<IReactlifecyclewpProps> {

    return (
      <div>
        <h1>ReactJS component's Lifecycle</h1>
        <h3>{this.state.stageTitle}</h3>
        <button onClick={this.updateState}>Clic Here To Update State Data!</button>
      </div>
    );
  }

  public componentWillUnmount() {
    console.log('Component will unmount has been called!');
  }
}

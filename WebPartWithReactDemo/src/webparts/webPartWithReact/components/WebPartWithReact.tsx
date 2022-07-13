import * as React from 'react';
import styles from './WebPartWithReact.module.scss';
import { IWebPartWithReactProps } from './IWebPartWithReactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ThemeProvider } from 'office-ui-fabric-react/lib/Foundation';

export default class WebPartWithReact extends React.Component<IWebPartWithReactProps, {}> {
  public render(): React.ReactElement<IWebPartWithReactProps> {
    const {
      description
    } = this.props;

    return (
      <section className={`${styles.webPartWithReact}`}>
        <div className={styles.welcome}>
          <h2>Well done!</h2>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <p>Datos aregados</p>
          <p className="${styles.description}">Absolute URL {escape(this.props.absoluteurl)}</p>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <p className='${ styles.description }'>Absolute URL {escape(this.props.absoluteurl)}</p>
          <p className='${ styles.description }'>Title {escape(this.props.sitetitle)}</p>
          <p className='${ styles.description }'>Relative URL {escape(this.props.relativeurl)}</p>
          <p className='${ styles.description }'>User Name {escape(this.props.username)}</p>
        </div>
      </section>
    );
  }
}

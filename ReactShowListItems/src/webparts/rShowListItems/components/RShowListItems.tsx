import * as React from 'react';
import styles from './RShowListItems.module.scss';
import { IRShowListItemsProps } from './IRShowListItemsProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jQuery from 'jquery';
import { ThemeProvider } from 'office-ui-fabric-react/lib/Foundation';

export interface IRShowListItemsWPState {
  listitems: [
    {
      "Title": "",
      "ID": "",
      "SoftwareName": ""
    }
  ]
}

export default class RShowListItems extends React.Component<IRShowListItemsProps, IRShowListItemsWPState> {
  
  static siteurl: string = "";

  public constructor (props: IRShowListItemsProps, state:IRShowListItemsWPState){
    super(props);
    this.state = {
      listitems: [
        {
          "Title": "",
          "ID": "",
          "SoftwareName": ""
        }
      ]
    };

    RShowListItems.siteurl = this.props.websiteurl;
  }

  public componentDidMount() {
    let reactcontexthandler = this;

    jQuery.ajax({
      url:`${RShowListItems.siteurl}/_api/web/lists/getbytitle('MicrosoftSoftware')/items`,
      type: "GET",
      headers: {'Accept': 'application/json; odata=verbose;'},
      success: function (resultData) {
        reactcontexthandler.setState({
          listitems: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
      }
    });
  }
  
  public render(): React.ReactElement<IRShowListItemsProps> {

    return (
      <div className={ styles.rShowListItems }>
        <table className={ styles.row }>
          {
            this.state.listitems.map(function(listitem, listitemkey) {
              let fullurl: string = `${RShowListItems.siteurl}/lists/MicrosoftSoftware/DispForm.aspx?ID=${listitem.ID}`;

              return (
                <tr>
                  <td>
                    <a className={styles.label} href={fullurl}>
                      {listitem.Title}
                    </a>
                  </td>
                  <td className={styles.label}>
                    {listitem.ID}
                  </td>
                  <td className={styles.label}>
                    {listitem.SoftwareName}
                  </td>
                </tr>
              );
            })
          }
        </table>

        <ol>
          {
            this.state.listitems.map(function (listitem, listitemkey) {
              let fullurl: string = `${RShowListItems.siteurl}/lists/MicrosoftSoftware/DispForm.aspx?ID=${listitem.ID}`;

              return (
                <li>
                  <a className={styles.label} href={fullurl}>
                    <span>{listitem.Title}</span> , <span>{listitem.ID}</span> , <span>{listitem.SoftwareName}</span>
                  </a>
                </li>
              );
              
            })
          }
        </ol>
      </div>
    );
  }
}

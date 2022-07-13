import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,
  //IPropertyPaneConfiguration 
} from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneWpWebPart.module.scss';
import * as strings from 'PropertyPaneWpWebPartStrings';
import { property } from 'lodash';

export interface IPropertyPaneWpWebPartProps {
  description: string;


  productname: string;
  productdescription: string;
  productcost: number;
  quantity: number;
  billamount: number;
  discount: number;
  netbillamount: number;

  currentTime: Date;
  IsCertified: boolean;
  Rating: number;
  processortype: string;
  InvoiceFileType: string;
  newProcessorType: string;
  discountCoupon: boolean;
}

export default class PropertyPaneWpWebPart extends BaseClientSideWebPart<IPropertyPaneWpWebPartProps> {

  //onInit, se utiliza como parametro de inicialización
  protected onInit(): Promise<void> {
    
    return new Promise<void>((resolve, _reject)=>{
      this.properties.productname = "Mouse";
      this.properties.productdescription = "Mouse Description";
      this.properties.quantity = 500;
      this.properties.productcost = 300;

      resolve(undefined);
    });
  }

  //Esta linea, permite que cuando se modifique la propiedad se vea reflejada en la pagina, por defaul es "false",
  //y para activar el boton se coloca true
  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneWp} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <table>
          <tr>
            <td>Product Name</td>
            <td>${ this.properties.productname }</td>
          </tr>
          <tr>
            <td>Description</td>
            <td>${ this.properties.productdescription}</td>
          </tr>
          <tr>
            <td>Product Cost</td>
            <td>${ this.properties.productcost }</td>
          </tr>
          <tr>
            <td>Product Quantity</td>
            <td>${ this.properties.quantity}</td>
          </tr>
          <tr>
            <td>Bill Amount</td>
            <td>${ this.properties.billamount = this.properties.productcost * this.properties.quantity }</td>
          </tr>
          <tr>
            <td>Discount</td>
            <td>${ this.properties.discount = this.properties.billamount * 10/100 }</td>
          </tr>
          <tr>
            <td>Net Bill Amount</td>
            <td>${ this.properties.netbillamount = this.properties.billamount - this.properties.discount }</td>
          </tr>
          <tr>
            <td>Is Certified?</td>
            <td>${ this.properties.IsCertified }</td>
          </tr>
          <tr>
            <td>Rating</td>
            <td>${ this.properties.Rating }</td>
          </tr>
          <tr>
            <td>Processor Type</td>
            <td>${ this.properties.processortype }</td>
          </tr>
          <tr>
            <td>File Type</td>
            <td>${ this.properties.InvoiceFileType }</td>
          </tr>
          <tr>
            <td>New Processor Type</td>
            <td>${ this.properties.newProcessorType }</td>
          </tr>
          <tr>
            <td>Do you have a discount coupon? </td>
            <td>${ this.properties.discountCoupon }</td>
          </tr>
        </table>
      </div>
    </section>`;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           //description: strings.PropertyPaneDescription
  //           description: "My Header description"
  //         },
  //         groups: [
  //           {
  //             //groupName: strings.BasicGroupName,
  //             groupName: "My groupName",
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 //label: strings.DescriptionFieldLabel
  //                 label: "My new label"
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return{
      pages:[
        {
          groups: [
            {
              groupName: "Product Details",
              groupFields: [

                PropertyPaneTextField('productname',{
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product name", "description": "Name property field"
                }),

                PropertyPaneTextField('productdescription',{
                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter Product Description", "description": "Name property field"
                }),

                PropertyPaneTextField('productcost',{
                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter Product Cost", "description": "Number property field",
                }),

                PropertyPaneTextField('quantity',{
                  label: "Product CQuantityost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter Product Quantity", "description": "Number property field"
                }),

                PropertyPaneToggle('IsCertified',{
                  key: 'IsCertified',
                  label: "Is it Certified?",
                  onText: 'ISI Certified!',
                  offText: 'Not an ISI Certified Product'
                }),

                PropertyPaneSlider('Rating',{
                  label: 'Select you Rating',
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 1
                }),

                PropertyPaneChoiceGroup('processortype',{
                  label: 'Choises',
                  options: [
                    {key: 'I5', text: 'Intel I5'},
                    {key: 'I7', text: 'Intel I7', checked: true},
                    {key: 'I9', text: 'Intel I9'}
                  ]
                }),

                PropertyPaneChoiceGroup('InvoiceFileType',{
                  label: 'Select Invoice File type:',
                  options: [
                    {
                      key: 'MSWord', text: 'MSWord',
                      imageSrc: 'https://apen.es/wp-content/uploads/2020/11/Microsoft-Word-Logotipo-2013-2019-300x169.jpg' ,
                      imageSize: {width: 32, height: 32},
                      selectedImageSrc: 'https://apen.es/wp-content/uploads/2020/11/Microsoft-Word-Logotipo-2013-2019-300x169.jpg'
                    },
                    {
                      key: 'MSExcel', text: 'MSExcel',
                      imageSrc: 'https://upload.wikimedia.org/wikipedia/commons/thumb/3/34/Microsoft_Office_Excel_%282019%E2%80%93present%29.svg/1200px-Microsoft_Office_Excel_%282019%E2%80%93present%29.svg.png' ,
                      imageSize: {width: 32, height: 32},
                      selectedImageSrc: 'https://upload.wikimedia.org/wikipedia/commons/thumb/3/34/Microsoft_Office_Excel_%282019%E2%80%93present%29.svg/1200px-Microsoft_Office_Excel_%282019%E2%80%93present%29.svg.png'
                    },
                    {
                      key: 'MSPowerPoint', text: 'MSPowerPoint',
                      imageSrc: 'https://norfipc.com/img/tutoriales/powerpoint.jpeg' ,
                      imageSize: {width: 32, height: 32},
                      selectedImageSrc: 'https://norfipc.com/img/tutoriales/powerpoint.jpeg'
                    },
                    {
                      key: 'OneNote', text: 'OneNote',
                      imageSrc: 'https://i0.wp.com/www.ardilu.com/wp-content/uploads/2020/06/Microsoft-OneNote-2013.png?ssl=1' ,
                      imageSize: {width: 32, height: 32},
                      selectedImageSrc: 'https://i0.wp.com/www.ardilu.com/wp-content/uploads/2020/06/Microsoft-OneNote-2013.png?ssl=1',
                      checked: true
                    }
                  ]
                }),

                PropertyPaneDropdown('newProcessorType',{
                  label: 'New Processor Type',
                  options: [
                    {key: 'I5', text: 'Intel I5'},
                    {key: 'I7', text: 'Intel I7'},
                    {key: 'I9', text: 'Intel I9'}
                  ],
                  selectedKey: 'I9'
                }),

                PropertyPaneCheckbox('discountCoupon', {
                  text: 'Do you have a Discount Coupon?',
                  checked: false,
                  disabled: false
                }),

                PropertyPaneLink('', {
                  href: 'https://www.amazon.com',
                  text: 'Buy Intel Processor from the best Seller',
                  target: '_blank',
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: 'Amazon'
                  }
                })
              ]
            }
          ]
        }
      ]
    }
  }
}

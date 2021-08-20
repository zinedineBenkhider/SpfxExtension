import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { sp } from '@pnp/sp/presets/all';
import { IFieldInfo } from "@pnp/sp/fields";
import { toast } from 'toast-notification-alert'
import DocumentListService, { Localisation } from '../service/DocumentListService';
import PropertyBagService from '../service/PropertyBagService';

//Sp context requirement
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');

export interface IUpdateDocPropertiesCommandSetProperties {
  siteUrl: string;
}

export interface IChoiceFieldInfo extends IFieldInfo {
  Choices: string[];
}

export default class UpdateDocPropertiesCommandSet extends BaseListViewCommandSet<IUpdateDocPropertiesCommandSetProperties> {
  @override
  public onInit(): Promise<void> {
    //init sp contecspfxContext
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  //Cette méthode s'éxécute dans le Context SP (SP.ClientContext)
  //Elle contient une fonctions et un CallBack.  
  private executeMDEcommand() {
    //init services
    let documentListService: DocumentListService = new DocumentListService(sp);
    let propertyBagService: PropertyBagService = new PropertyBagService();
    
    //CallBack qui s'éxécute quand les Properties Bag sont récupérées
    let onGetLastUpdateDateTimeSuccess=(sender: any, args: SP.ClientRequestSucceededEventArgs) => {
      let lastUpdateDateTime = propertyBagService.getPropertyLastUpdateDateTime();
      //Mettre à jour les éléments de la liste de Documents
      updateListItems(lastUpdateDateTime);
    }

    //Récupérer les propriétés de RootFolder de la liste documents 
    propertyBagService.getProperties(onGetLastUpdateDateTimeSuccess);
    
    //Mofifier les colonnes de la liste de documents
    let updateListItems=async lastUpdateDateTime=>{
        let validLanguages;
        let validTypesDoc = [];
        let validLocalisations: Localisation[] = [];
        //Récupérer les éléments du Set Localisation
        let getLocalisationsPromise = documentListService.getLocalisationsTermeStore().then(items => {
          let element;
          for (let i = 0; i < items.length; i++) {
            element = items[i];
            validLocalisations.push(new Localisation(element.labels[0].name, element.id));
          }
        });
        //récupérer la choice list du champs TypeDoc
        let getChoicesPromise = documentListService.getChoicesTypeDoc().then((fieldData: IChoiceFieldInfo) => {
          validTypesDoc = fieldData.Choices;
        });

        //récupérer la list Langues du champs Langue
        let getLanguagesListPromise = documentListService.getLanguages().then(langues => {
          validLanguages = langues;
        });

        //Attendre que les Promises terminent. Nb: on attend que la plus lente des 3
        await getLocalisationsPromise;
        await getChoicesPromise;
        await getLanguagesListPromise;

        documentListService.initDocItems();
        //Sélectionner les fichiers modifiés depuis la dérniére mise à jour, et modifié leurs champs si le nom du fichier est de bon format.
        documentListService.getFilesModified(lastUpdateDateTime).then(response => {
          if (response.length == 0) {
            toast.show({ title: '', message: "Toutes les informations sont à jour", type: "info", newestOnTop: false });
          }
          else {
            response.forEach((element, index, elements) => {
              let fileName = element.File.Name;
              let message = "";
              let fileNameSplit = fileName.split("_");
              if (fileNameSplit.length == 6) {
                //Extraction des propriétés depuis le nom du fichier
                let typeDoc = fileNameSplit[1];
                let language = fileNameSplit[2];
                let localisation = fileNameSplit[3];
                let subject = fileNameSplit[4];
                let datePub = fileNameSplit[5].split(".")[0];

                //Vérifier la validité des propriétés extraites
                let languageId = DocumentListService.isValidLanguage(language, validLanguages);
                let validTypeDoc = DocumentListService.isValidTypeDoc(typeDoc, validTypesDoc);
                let validLocalisationId = DocumentListService.isValidLocalisation(localisation, validLocalisations);
                let validDateFormat = DocumentListService.isValidDateFormat(datePub);

                //Créer les requetes update des propriétés dans un Batch(File d'attente)
                documentListService.initBatch();
                //Update Locaisation Field
                if (validLocalisationId != -1) {
                  documentListService.updateValueOfLocalisation(element.Id, localisation, validLocalisationId);
                }
                else {
                  message = message + "Localisation";
                }
                //Update Langue Field
                if (languageId != -1) {
                  documentListService.updatevalueOfLanguage(element.Id, languageId);
                }
                else {
                  let m = message == "" ? "Langue" : ", Langue";
                  message = message + m;
                }
                //Update TypeDoc Field
                if (validTypeDoc) {
                  documentListService.updateValueOfTypeDoc(element.Id, typeDoc);
                }
                else {
                  let m = message == "" ? "TypeDoc" : ", TypeDoc";
                  message = message + m;
                }
                //Update DatePublication Field
                if (validDateFormat != "") {
                  documentListService.updateValueOfPublicationDate(element.Id, validDateFormat);
                }
                else {
                  let m = message == "" ? "PublicationDate" : ", PublicationDate";
                  message = message + m;
                }
                //Update Subject Field
                documentListService.updateValueOfSubject(element.Id, subject);
                if (message != "") {
                  if (message.split(",").length >= 2) {
                    message = "Les propriétés: " + " " + message + " du fichier " + fileName + " ne sont pas misent à jour";
                  }
                  else {
                    message = "La propriété: " + " " + message + " du fichier " + fileName + " n'est pas mise à jour";
                  }
                  toast.show({ title: '', message: message, type: "info", newestOnTop: false });
                }
                //éxécuter les requtes update du Batch
                documentListService.executeBatch().then(() => {
                  //Quand toutes les propriétés du dernier élement de la liste sont modifiées on met à jour la propriété LastUpdateDateTime avec la date et l'heure du moment (now)
                  if (index === elements.length - 1) {
                    propertyBagService.setPropertyLastUpdateDateTime();
                  }
                });
              }
            }
            );
          }
        }
        )
    }
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
  }

  /*Cette méthode se déclenche au moment de l'éxécution d'une commande*/
  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'MDE':
        //éxécuter la fonction updateItems dans le ccontext SP
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', this.executeMDEcommand);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}


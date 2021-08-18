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
import UpdateDocPropertiesService from '../service/UpdatePropertiesService';

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

export class Localisation {
  name: string;
  id: any;
  constructor(name: string, id: any) {
    this.name = name;
    this.id = id;
  }
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
  // Elle contient des fonctions et des CallBack.  
  private executeCommandMDE() {
    let clientContext: SP.ClientContext;
    let rootFolderProperties: SP.PropertyValues;
    let rootFolder: SP.Folder;
    //CallBack qui se déclanche quand les propriétés de RootFolder sont récupérés, ou si la récupération échoue (LastUpdateDateTime n'éxiste pas encore). 
    let onGetRootFolderPropreties = (sender: any, args: SP.ClientRequestSucceededEventArgs) => {
      let lastUpdateDateTime;
      try {
        lastUpdateDateTime = rootFolderProperties.get_item("LastUpdateDateTime");
      } catch (error) {//La propriété LastUpdateDateTime n'éxiste pas, on prend une date antérieur. Aprés la mise à jour de la liste LastUpdateDateTime sera créé avec cette date. 
        lastUpdateDateTime = dateToIsoString(new Date('01 January 1900 00:00 UTC'));
      }
      updateItems(lastUpdateDateTime);
    };

    //récupérer les properiétés de root folder de la liste de Documents
    //UPDATE_HERE
    clientContext = new SP.ClientContext("https://zinedinebenkhider.sharepoint.com/sites/MetaDataEnterprise");
    let oweb = clientContext.get_web();
    clientContext.load(oweb);
    rootFolder = oweb.get_lists().getByTitle("Documents").get_rootFolder();
    rootFolderProperties = rootFolder.get_properties();
    clientContext.load(rootFolderProperties);
    clientContext.executeQueryAsync(onGetRootFolderPropreties, onGetRootFolderPropreties);

    //Cette Fonction met à jour les éléments de la liste de Documents
    let updateItems = async (lastUpdateDateTime) => {
      let validLanguages;
      let validTypesDoc = [];
      let validLocalisations: Localisation[] = [];

      //UPDATE_HERE : ID du groupe CCI termeStore => 45f13976-c5f0-4f49-b7cf-004afa72d7b4 | ID du Set Localisation=> 7ee71116-9a06-40fe-be85-b66ee794847d
      //Récupérer les élément du Set Localisation
      let getLocationsPromise = sp.termStore.groups.getById("45f13976-c5f0-4f49-b7cf-004afa72d7b4").sets.getById("7ee71116-9a06-40fe-be85-b66ee794847d").children().then(items => {
        let element;
        for (let i = 0; i < items.length; i++) {
          element = items[i];
          validLocalisations.push(new Localisation(element.labels[0].name, element.id));
        }
      });

      //récupérer la choice list du champs TypeDoc
      let getChoicesPromise = sp.web.lists.getByTitle('Documents').fields.getByInternalNameOrTitle('TypeDoc').select('Choices').get().then((fieldData: IChoiceFieldInfo) => {
        validTypesDoc = fieldData.Choices;
      });

      //récupérer la list Langues du champs Langue
      let getLanguagesListPromise = sp.web.lists.getByTitle('Langues').items.select('Title', 'Id').get().then(langues => {
        validLanguages = langues;
      });

      //Attendre que les promises terminent. nb: on attend que la plus lente des 3
      await getLocationsPromise;
      await getChoicesPromise;
      await getLanguagesListPromise;
      let items = sp.web.lists.getByTitle('Documents').items;

      //Sélectionner les fichiers modifiés depuis la dérniére mise à jour, et modifié leurs champs si le nom du fichier est de bon format.
      //FSObjType ne 1 => ne pas prendre les dossiers
      //Modified ge datetime'${lastUpdateDateTime}' =>  champs Modified >lastUpdateDateTime
      items.select('Id', 'File/Name').expand('File/Name').
        filter(`FSObjType ne 1 and Modified ge datetime'${lastUpdateDateTime}'`)
        .get().then(response => {
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
                let languageId = isValidLanguage(language, validLanguages);
                let validTypeDoc = isValidTypeDoc(typeDoc, validTypesDoc);
                let validLocalisationId = isValidLocalisation(localisation, validLocalisations);
                let validDateFormat = isValidDateFormat(datePub);

                //Créer les requetes update des propriétés dans un Batch(File d'attente)
                const batch = sp.web.createBatch();
                //Update Locaisation Field
                if (validLocalisationId != -1) {
                  items.getById(element.Id).inBatch(batch).validateUpdateListItem([
                    { FieldName: 'Localisation', FieldValue: localisation + '|' + validLocalisationId + ';' },
                  ]).then(result => {
                    console.log(JSON.stringify(result));
                  });
                }
                else {
                  message = message + "Localisation";
                }
                //Update Langue Field
                if (languageId != -1) {
                  items.getById(element.Id).inBatch(batch).update({
                    LangueId: languageId,
                  });
                }
                else {
                  let m = message == "" ? "Langue" : ", Langue";
                  message = message + m;
                }
                //Update TypeDoc Field
                if (validTypeDoc) {
                  items.getById(element.Id).inBatch(batch).update({
                    TypeDoc: typeDoc,
                  });
                }
                else {
                  let m = message == "" ? "TypeDoc" : ", TypeDoc";
                  message = message + m;
                }
                //Update DatePublication Field
                if (validDateFormat != "") {
                  items.getById(element.Id).inBatch(batch).update({
                    DatePublication: validDateFormat,
                  });
                }
                else {
                  let m = message == "" ? "PublicationDate" : ", PublicationDate";
                  message = message + m;
                }
                if (message != "") {
                  if (message.split(",").length >= 2) {
                    message = "Les propriétés: " + " " + message + " du fichier " + fileName + " ne sont pas misent à jour";
                  }
                  else {
                    message = "La propriété: " + " " + message + " du fichier " + fileName + " n'est pas mise à jour";
                  }
                  toast.show({ title: '', message: message, type: "info", newestOnTop: false });
                }
                items.getById(element.Id).inBatch(batch).update({
                  Subject: subject,
                });
                /* Mettre à jour les 4 champs au mm temps. Si la modification d'un seule champs échoue, c'est toutes les modifs qui échouent.
                items.getById(element.Id).inBatch(batch).update({
                  TypeDoc: typeDoc,
                  Subject: subject,
                  DatePublication: datePublication,
                  LangueId: languageId,
                }).then(result => {
                  console.log(JSON.stringify(result));
                });
                */
                batch.execute().then(() => {
                  //Quand toutes les propriétés du dernier élement de la liste sont modifiées
                  //On met à jour la propriété LastUpdateDateTime avec la date et l'heure du moment (now)
                  if (index === elements.length - 1) {
                    let nowDateTime = dateToIsoString(new Date());
                    rootFolder.get_properties().set_item("LastUpdateDateTime", nowDateTime);
                    rootFolder.update();
                    clientContext.executeQueryAsync();
                  }
                });
              }
            }
            );
          }

        }
        )
    }
    //Conversion d'une date vers le format ISO sans millisecondes
    let dateToIsoString = (date: Date) => {
      return date.toISOString().split('.')[0] + "Z";//le split sert à éliminer les millisecondes 
    }

    //Cette fonction Vérifie si une langue est présente dans une liste de Langues 
    //valideLanguages : list d'objet Langue qui contient deux attributs: Title, Id
    //@return: l'Id si la valeur est présente, sinon -1 
    let isValidLanguage = (language: string, validLanguages) => {
      let languageId = -1;
      validLanguages.forEach(element => {
        if (element.Title == language) {
          languageId = element.Id;
        }
      });
      return languageId;
    }

    //Cette fonction vérifie si une localisation est présente dans une liste de localisations 
    //validLocalisations : list d'objet Localisation qui contient deux attributs: name, id
    //@return: l'Id si la valeur est présente, sinon -1 
    let isValidLocalisation = (localisationName: string, validLocalisations: Localisation[]) => {
      let localisationId = -1;
      for (let index = 0; index < validLocalisations.length; index++) {
        if (validLocalisations[index].name == localisationName) {
          localisationId = validLocalisations[index].id;
        }
      }
      return localisationId;
    }

    //Cette fonction vérifie si un TypeDoc est présente dans une liste de TypeDoc 
    //validTypesDoc : list de TypeDoc
    //@return: true si la valeur est présente, sinon false 
    let isValidTypeDoc = (value: string, validTypesDoc: string[]) => {
      for (let index = 0; index < validTypesDoc.length; index++) {
        if (validTypesDoc[index] == value) {
          return true
        }
      }
      return false;
    }

    //Cette fonction vérifie si une date est une chaine de 6 charactere
    //validLocalisations : list de TypeDoc
    //@return: la date sous format ISO si la date est valide, sinon une chaine vide
    let isValidDateFormat = (datePub) => {
      let year = datePub.substring(0, 2);
      let month = datePub.substring(2, 4);
      let day = datePub.substring(4);
      if (datePub.length == 6 && month <= 12 && day <= 31) {
        return "20" + year + "-" + month + "-" + day + "T23:00:00Z";  // ex: iso date format "2019-05-24T23:00:00Z"
      }
      return "";
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
        //éxécuter la fonction executeCommandMDE dans le ccontext SP
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', this.executeCommandMDE);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}


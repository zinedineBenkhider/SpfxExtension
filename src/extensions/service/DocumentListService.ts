import { IItems, SPBatch, SPRest } from '@pnp/sp/presets/all';
export class Localisation {
  public name: string;
  public id: any;
  constructor(name: string, id: any) {
    this.name = name;
    this.id = id;
  }
}

export default class DocumentListService {
  private sp: SPRest;
  private batch: SPBatch;
  private docItems: IItems;

  constructor(sp: SPRest) {
    this.sp = sp;
  }

  public initDocItems() {
    this.docItems = this.sp.web.lists.getByTitle('Documents').items;
  }

  public initBatch() {
    this.batch = this.sp.web.createBatch();
  }
  public executeBatch(): Promise<void> {
    return this.batch.execute();
  }
  public getLanguages(): Promise<any> {
    return this.sp.web.lists.getByTitle('Langues').items.select('Title', 'Id').get();
  }

  public getChoicesTypeDoc(): Promise<any> {
    return this.sp.web.lists.getByTitle('Documents').fields.getByInternalNameOrTitle('TypeDoc').select('Choices').get();
  }
  public getLocalisationsTermeStore(): Promise<any> {
    //cc1a4b2e-d36c-4861-b9ba-1a01224019f4 6a921481-3b68-4b58-acea-60893f914f98
    //UPDATE_HERE : ID du groupe CCI termeStore => 45f13976-c5f0-4f49-b7cf-004afa72d7b4 | ID du Set Localisation=> 7ee71116-9a06-40fe-be85-b66ee794847d
    return this.sp.termStore.groups.getById("4212d6d5-226e-4798-8ac4-d47d07d34d33").sets.getById("12d8cbd3-3c90-4233-abca-7791333730a3").children();
  }

  //FSObjType ne 1 => ne pas prendre les dossiers
  //Modified ge datetime'${lastUpdateDateTime}' =>  champs Modified >lastUpdateDateTime
  public getFilesModified(lastUpdateDateTime): Promise<any> {
    return this.docItems.select('Id','Modified', 'File/Name').expand('File/Name').
      filter(`FSObjType ne 1 and Modified ge datetime'${lastUpdateDateTime}'`)
      .get();
  }

  public updateValueOfLocalisation(id, localisation, validLocalisationId) {
    this.docItems.getById(id).inBatch(this.batch).validateUpdateListItem([
      { FieldName: 'Localisation', FieldValue: localisation + '|' + validLocalisationId + ';' },
    ]);
  }

  public updatevalueOfLanguage(id, value) {
    this.docItems.getById(id).inBatch(this.batch).update({
      LangueId: value,
    });
  }

  public updateValueOfTypeDoc(id, value) {
    this.docItems.getById(id).inBatch(this.batch).update({
      TypeDoc: value,
    });
  }

  public updateValueOfSubject(id, value) {
    this.docItems.getById(id).inBatch(this.batch).update({
      Subject: value,
    });
  }
  public updateValueOfPublicationDate(id, value) {
    this.docItems.getById(id).inBatch(this.batch).update({
      DatePublication: value,
    });
  }

  //Cette fonction V??rifie si une langue est pr??sente dans une liste de Langues 
  //valideLanguages : list d'objet Langue qui contient deux attributs: Title, Id
  //@return: l'Id si la valeur est pr??sente, sinon -1 
  public static isValidLanguage(language: string, validLanguages) {
    let languageId = -1;
    validLanguages.forEach(element => {
      if (element.Title == language) {
        languageId = element.Id;
      }
    });
    return languageId;
  }

  //Cette fonction v??rifie si une localisation est pr??sente dans une liste de localisations 
  //validLocalisations : list d'objet Localisation qui contient deux attributs: name, id
  //@return: l'Id si la valeur est pr??sente, sinon -1 
  public static isValidLocalisation(localisationName: string, validLocalisations: Localisation[]) {
    let localisationId = -1;
    for (let index = 0; index < validLocalisations.length; index++) {
      if (validLocalisations[index].name == localisationName) {
        localisationId = validLocalisations[index].id;
      }
    }
    return localisationId;
  }

  //Cette fonction v??rifie si un TypeDoc est pr??sente dans une liste de TypeDoc 
  //validTypesDoc : list de TypeDoc
  //@return: true si la valeur est pr??sente, sinon false 
  public static isValidTypeDoc(value: string, validTypesDoc: string[]) {
    for (let index = 0; index < validTypesDoc.length; index++) {
      if (validTypesDoc[index] == value) {
        return true;
      }
    }
    return false;
  }

  //Cette fonction v??rifie si une date est une chaine de 6 charactere
  //validLocalisations : list de TypeDoc
  //@return: la date sous format ISO si la date est valide, sinon une chaine vide
  public static isValidDateFormat(datePub) {
    let year = datePub.substring(0, 2);
    let month = datePub.substring(2, 4);
    let day = datePub.substring(4);
    if (datePub.length == 6 && month <= 12 && month > 0 && day <= 31 && day > 0) {
      return "20" + year + "-" + month + "-" + day + "T23:00:00Z";  // ex: iso date format "2019-05-24T23:00:00Z"
    }
    return "";
  }
}
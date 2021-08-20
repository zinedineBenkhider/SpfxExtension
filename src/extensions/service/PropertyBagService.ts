export default class PropertyBagService {
    private clientContext: SP.ClientContext;
    private rootFolderProperties: SP.PropertyValues;
    private rootFolder: SP.Folder;
    constructor(){
        //UPDATE_HERE
        this.clientContext = new SP.ClientContext("https://zinedinebenkhider.sharepoint.com/sites/MetaDataEnterprise");
        let oweb = this.clientContext.get_web();
        this.clientContext.load(oweb);
        this.rootFolder = oweb.get_lists().getByTitle("Documents").get_rootFolder();
    }

    public getProperties(callBack:(sender: any, args: SP.ClientRequestSucceededEventArgs) =>void){
        this.rootFolderProperties = this.rootFolder.get_properties();
        this.clientContext.load(this.rootFolderProperties);
        this.clientContext.executeQueryAsync(callBack, callBack);
    }

    public getPropertyLastUpdateDateTime(){
        let lastUpdateDateTime;
        try{
            lastUpdateDateTime= this.rootFolderProperties.get_item("LastUpdateDateTime");
          }
          catch (error){
            lastUpdateDateTime = PropertyBagService.dateToIsoString(new Date('01 January 1900 00:00 UTC'));
          }
        return lastUpdateDateTime;
    }

    public setPropertyLastUpdateDateTime(){
        let nowDateTime = PropertyBagService.dateToIsoString(new Date());
        this.rootFolder.get_properties().set_item("LastUpdateDateTime", nowDateTime);
        this.rootFolder.update();
        this.clientContext.executeQueryAsync();
    }


 //Conversion d'une date vers le format ISO sans millisecondes
  static dateToIsoString  (date: Date) {
    return date.toISOString().split('.')[0] + "Z";//le split sert à éliminer les millisecondes 
  }
}
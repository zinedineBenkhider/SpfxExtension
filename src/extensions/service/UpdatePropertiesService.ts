import { SPRest } from '@pnp/sp/presets/all';

export default class UpdateDocPropertiesService {
    static getLanguages(sp:SPRest):Promise<any> {
        return  sp.web.lists.getByTitle('Langues').items.select('Title', 'Id').get();
    }

    
   
}
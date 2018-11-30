
import { sp,  Item, ItemAddResult, ItemUpdateResult } from "sp-pnp-js";
import CONSTANTS from "./Constants"

export default class CommentsServiceService {
    public getCurrentUser() {
        return sp.web.currentUser.configure(CONSTANTS.HeaderNoMedata)
            .select("Id","Title", "Groups/Title","Groups/Id", "LoginName","FirstName","EMail")
            .expand("Groups").get();
    }

    public getData(listName:string)
    {
        return sp.web.lists.getByTitle(listName).configure(CONSTANTS.HeaderNoMedata).items
            .select("ID","Title","DFPostDescription","Created","Author/Title"
                    ,"Author/EMail","Author/Office","TotalComments","TotalLikes"
                    ,"DFCategory/Title","DFCategory/ID","DFLocation/Title","DFLocation/ID"
                    ,"Audience/Title","Audience/ID","LikesCount","LikedBy/ID")
            .expand("Author","DFCategory","DFLocation","Audience","LikedBy")
            .filter("DFStatus eq 'Submitted'")
            .get(); 

    }

    public getAllProjects()
    {
        return sp.web.lists.getByTitle(CONSTANTS.ProjectList).configure(CONSTANTS.HeaderNoMedata).items
        .select("ID","Title","DFPostDescription","Created","Author/Title","Author/EMail","Author/Office","TotalComments"
                ,"TotalLikes","DFCategory/Title","DFCategory/ID","DFLocation/Title","DFLocation/ID","Audience/Title"
                ,"Audience/ID","LikesCount","LikedBy/ID")
        .expand("Author","DFCategory","DFLocation","Audience","LikedBy")
        .filter("DFStatus eq 'Submitted'")
        .get(); 
    } 

    public getAllQuestions(){
        return sp.web.lists.getByTitle(CONSTANTS.QuestionList)
            .configure(CONSTANTS.HeaderNoMedata).items
            .select("ID","Title","DFPostDescription","Created","Author/Title","Author/EMail","Author/Office","TotalComments"
            ,"TotalLikes","DFCategory/Title","DFCategory/ID","DFLocation/Title","DFLocation/ID","Audience/Title","Audience/ID","LikesCount","LikedBy/ID")
            .expand("Author","DFCategory","DFLocation","Audience","LikedBy")
            .filter("DFStatus eq 'Submitted'")
            .get()
            .catch((e)=>{}); 
    }

    public getConfigurationdata(listName:string){
        return sp.web.lists.getByTitle(listName)
            .configure(CONSTANTS.HeaderNoMedata).items
            .select("ID","Title")
            .get()
            .catch((e)=> {}) 
    }

    public updateListItem(listTitle:string,itemId:number,itemInformation:any)
    {
        let list =sp.web.lists.getByTitle(listTitle);
        list.items.getById(itemId).update(itemInformation).then(i=> {console.log(i);})
        .catch((error)=>{console.log(error)})
    }

    public getListItemById(listTitle:string,itemId:number,)
    {
        return sp.web.lists.getByTitle(listTitle).items
            .getById(itemId)
            .select("ID","LikesCount","LikedBy/ID","FavouritesAssociates")
            .expand("LikedBy")
            .get()
    }

    public getCoverImages(listName:string)
    {
        return sp.web.lists.getByTitle(listName).configure(CONSTANTS.HeaderNoMedata).items.get(); 
    } 

    public GetConfigurations(component:string){
        return sp.web.lists.getByTitle(CONSTANTS.ConfigurationList).configure(CONSTANTS.HeaderNoMedata).items
        .select("Title","Value").get()
    }

    public getListDetails(listName:string) {
        return sp.web.lists.getByTitle(listName).items.select("Title","Id").get();
    }

    public getItemDetails(itemID:number, listname:string, select:string, expand:string) {
        return sp.web.lists.getByTitle(listname).items.getById(itemID)
        .select(select)
        .expand(expand)
        .get()
    }

    public updateItem(itemID:number, data: {},listname: string) {

    return sp.web.lists.getByTitle(listname).configure(CONSTANTS.HeaderNoMedata).items.getById(itemID)

    .update(data);

    }

    public addItem(data: {},
    listname: string) {

    return sp.web.lists.getByTitle(listname).configure(CONSTANTS.HeaderNoMedata).items

    .add(data);

    }


    public getLoggedInUserDetails() {
        return sp.profiles.myProperties.get();
    }
    public DeleteAttachment (fileName:string) {
        return sp.web.getFolderByServerRelativeUrl(CONSTANTS.FolderRelativePath).files.getByName(fileName).delete()
    }
    public GetFileIcon (fileName:string,size:number) {
        return sp.web.mapToIcon(fileName,size,"");
    }

    public getItemDetailsFilterBased(listname:string, select:string,filter:string,expand:string) {
        return sp.web.lists.getByTitle(listname).items
            .select(select)
            .filter(filter)
            .expand(expand)
            .orderBy("Created",false)
            .get()
    }

    public DeleteItem (listName:string,itemId: number) {
        return sp.web.lists.getByTitle(listName).items.getById(itemId).delete()
    }
}




import { sp,  Item, ItemAddResult, ItemUpdateResult } from "sp-pnp-js";
import CONSTANTS from "./Constants"
import LogManager from '../../LogManager';

export default class CommentsServiceService {
    public getCurrentUser() {
        return sp.web.currentUser.configure(CONSTANTS.HeaderNoMedata)
            .select("Id","Title", "Groups/Title","Groups/Id", "LoginName","FirstName","EMail")
            .expand("Groups").get()
            .catch((error) => {
                 LogManager.logException(error
                    ,"Error occured while fetching getting current user details"
                    ,"CommentsService"
                    ,"getCurrentUser");
            });
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
            .get()
            .catch((error) => {
                 LogManager.logException(error
                    ,"Error occured while fetching data from list"
                    ,"CommentsService"
                    ,"getData");
            });

    }

    public getAllProjects()
    {
        return sp.web.lists.getByTitle(CONSTANTS.ProjectList).configure(CONSTANTS.HeaderNoMedata).items
        .select("ID","Title","DFPostDescription","Created","Author/Title","Author/EMail","Author/Office","TotalComments"
                ,"TotalLikes","DFCategory/Title","DFCategory/ID","DFLocation/Title","DFLocation/ID","Audience/Title"
                ,"Audience/ID","LikesCount","LikedBy/ID")
        .expand("Author","DFCategory","DFLocation","Audience","LikedBy")
        .filter("DFStatus eq 'Submitted'")
        .get()
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching all projects"
                ,"CommentsService"
                ,"getAllProjects");
        });
    } 

    public getAllQuestions(){
        return sp.web.lists.getByTitle(CONSTANTS.QuestionList)
            .configure(CONSTANTS.HeaderNoMedata).items
            .select("ID","Title","DFPostDescription","Created","Author/Title","Author/EMail","Author/Office","TotalComments"
            ,"TotalLikes","DFCategory/Title","DFCategory/ID","DFLocation/Title","DFLocation/ID","Audience/Title","Audience/ID","LikesCount","LikedBy/ID")
            .expand("Author","DFCategory","DFLocation","Audience","LikedBy")
            .filter("DFStatus eq 'Submitted'")
            .get()
            .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching all questions"
                ,"CommentsService"
                ,"getAllProjects");
            });
    }

    public getConfigurationdata(listName:string){
        return sp.web.lists.getByTitle(listName)
            .configure(CONSTANTS.HeaderNoMedata).items
            .select("ID","Title")
            .get()
            .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching configuration data"
                ,"CommentsService"
                ,"getConfigurationdata");
            });
    }

    public updateListItem(listTitle:string,itemId:number,itemInformation:any)
    {
        let list =sp.web.lists.getByTitle(listTitle);
        list.items.getById(itemId).update(itemInformation).then(i=> {console.log(i);})
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while updating list item"
                ,"CommentsService"
                ,"updateListItem");
            });
    }

    public getListItemById(listTitle:string,itemId:number,)
    {
        return sp.web.lists.getByTitle(listTitle).items
            .getById(itemId)
            .select("ID","LikesCount","LikedBy/ID","FavouritesAssociates")
            .expand("LikedBy")
            .get()
             .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching list item by Id"
                ,"CommentsService"
                ,"getListItemById");
            });
    }

    public getCoverImages(listName:string)
    {
        return sp.web.lists
            .getByTitle(listName)
            .configure(CONSTANTS.HeaderNoMedata)
            .items
            .get()
            .catch((error) => {
                    LogManager.logException(error
                    ,"Error occured while fetching cover images"
                    ,"CommentsService"
                    ,"getCoverImages");
            }); 
    } 

    public GetConfigurations(component:string){
        return sp.web.lists
            .getByTitle(CONSTANTS.ConfigurationList)
            .configure(CONSTANTS.HeaderNoMedata)
            .items
            .select("Title","Value")
            .get()
            .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching configuration data"
                ,"CommentsService"
                ,"GetConfigurations");
            }); 
    }

    public getListDetails(listName:string) {
        return sp.web.lists
        .getByTitle(listName)
        .items
        .select("Title","Id")
        .get()
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching basic list details"
                ,"CommentsService"
                ,"getListDetails");
            }); 
    }

    public getItemDetails(itemID:number, listname:string, select:string, expand:string) {
        return sp.web.lists.getByTitle(listname).items.getById(itemID)
        .select(select)
        .expand(expand)
        .get()
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching item details"
                ,"CommentsService"
                ,"getItemDetails");
            }); 
    }

    public updateItem(itemID:number, data: {},listname: string) {
        return sp.web.lists.getByTitle(listname).configure(CONSTANTS.HeaderNoMedata).items.getById(itemID)
        .update(data)
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while updating item"
                ,"CommentsService"
                ,"updateItem");
            }); 

    }

    public addItem(data: {},
        listname: string) {
        return sp.web.lists.getByTitle(listname).configure(CONSTANTS.HeaderNoMedata).items
        .add(data)
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while adding item"
                ,"CommentsService"
                ,"addItem");
            }); 

    }


    public getLoggedInUserDetails() {
        return sp.profiles.myProperties.get()
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fetching logged in user details"
                ,"CommentsService"
                ,"getLoggedInUserDetails");
            }); 
    }
    public DeleteAttachment (fileName:string) {
        return sp.web.getFolderByServerRelativeUrl(CONSTANTS.FolderRelativePath)
            .files
            .getByName(fileName)
            .delete()
            .catch((error) => {
                LogManager.logException(error
                ,"Error occured while deleting the attachment"
                ,"CommentsService"
                ,"DeleteAttachment");
            }); 
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
            .catch((error) => {
                LogManager.logException(error
                ,"Error occured while fecthing filtered items"
                ,"CommentsService"
                ,"getItemDetailsFilterBased");
            }); 
    }

    public DeleteItem (listName:string,itemId: number) {
        return sp.web.lists.getByTitle(listName).items.getById(itemId)
        .delete()
        .catch((error) => {
                LogManager.logException(error
                ,"Error occured while deleting an item"
                ,"CommentsService"
                ,"DeleteItem");
            }); 
    }
}



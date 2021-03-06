
/*** This file is used to render comments and reply in the details page**/

import * as React from 'react';
import styles from './Comments.module.scss';
import { ICommentsProps } from './ICommentsProps';
// import CONSTANTS from "../common/constants";
import service from "./CommentsService";
import { Persona, PersonaSize, PersonaPresence } from  "office-ui-fabric-react/lib/Persona";
import { Button } from  "office-ui-fabric-react/lib/Button";
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import LogManager from '../../LogManager';

// export interface ICommentsProps {
//     ParentMode: string;
// }
export interface ICommentsState {
    comments: string;
    replyToComments: string;
    addedComments: any[];
    itemID: string;
    displayReplyBlock: string;
    errorReply: boolean
    errorPost: boolean;
}

export default class CommentReplySection extends React.Component<ICommentsProps,ICommentsState>{
    private service =new service();

    constructor(props:ICommentsProps, state: ICommentsState) {
        super(props);
        this.state = {
            comments: '',
            replyToComments: '',
            addedComments: [],
            itemID: '',
            displayReplyBlock: '',
            errorReply: false,
            errorPost: false
        }
        this.addCommentReply = this.addCommentReply.bind(this);
    }

    public componentDidMount() {

        //get all the comments

        this.getCommentsDetails();

    }

    public componentWillMount() {
        this.GetItemID();
    }

    public GetItemID() {
       // let id = GetUrlKeyValue('ItemID');
        var queryParameters = new UrlQueryParameterCollection(window.location.href);
        let id= queryParameters.getValue(this.props.queryStrItemIdFieldName);
        //id="55";
        this.setState({
            itemID: id
        });
    }

    private getCommentsDetails() {
        let strFilter:string = "ParentItemId eq '" + this.state.itemID + "'";
        let strExpand:string = "Author";
        let strSelect:string = 'ID,ParentCommentId,CommentBody,Created,Author/ID,Author/Title,Author/Office,Author/EMail,Author/UserName';
        let filteredComments:any[] = [];
        let listName:string = this.props.listName;
        if (this.state.itemID !=null && listName !='') {
            this.service.getItemDetailsFilterBased(listName,strSelect, strFilter,strExpand).then((result:any) => {
                //set to complete state and username in expand
                if (result !=undefined && result.length >0) {
                    result.map((value,index) => {
                        let objFinal:any = { objParent: {}, objReplies: [] };

                        if (value.ParentCommentId == null || value.ParentCommentId =='') {
                            objFinal.objParent =value;
                            result.map((reply,key) => {
                                if (reply.ParentCommentId ==value.ID) {
                                objFinal.objReplies.push(reply);
                                }
                            });
                            filteredComments.push(objFinal);
                        }
                    });
                    console.log("final comment list::");
                    console.log(filteredComments);
                    this.setState({
                        addedComments: filteredComments
                    });
                }
                else {
                    this.setState({
                        addedComments: []
                    });
                }
            }).catch((error) => {
                 LogManager.logException(error
                    ,"Error occured while fetching comments for the item"
                    ,"Comments"
                    ,"getCommentsDetails");
               
                this.setState({
                    addedComments: []
                });
            })
        }
    }

    private addCommentReply(id:string) {
        let commentMode:
        string = id.split('#')[0];
        if (commentMode =="Reply" && this.state.replyToComments =='') {
            this.setState({
                errorReply: true
            });
            return false;
        }

        else if (commentMode != "Reply" && this.state.comments =='') {
            this.setState({
                errorPost: true
            });
            return false;
        }
        else {
            var item:
            any = {};
            let listName:string = this.props.listName;
            item = {
                ParentItemId: this.state.itemID,
                CommentBody: commentMode != "Reply" ? 
                this.state.comments :this.state.replyToComments,
                ParentCommentId: commentMode == "Reply" ? id.split('#')[1] :''
            }

            this.service.addItem(item, listName).then((data:any) => {
                this.getCommentsDetails();
                this.setState({
                    displayReplyBlock: '',
                    comments: '',
                    replyToComments: ''
                });
            });
            return true;
        }
    }

    private setDisplayPost(id:string) {
        id = this.state.displayReplyBlock == id ? '' : id;
        this.setState({
            displayReplyBlock: id,
            errorReply: false
        })
    }
    //Change Control Events
    private onChangeControls = (event:any) => {
        var state = this.state;
        state[event.target.name] =event.target.value;
        this.setState(state);
        this.setState({
            errorPost: false,
            errorReply: false
        });
    }

    public render(): React.ReactElement<ICommentsProps> {
        let dateformate = {
            month: 'long',
            year: 'numeric',
            day: '2-digit',
            hour: 'numeric',
            minute: 'numeric'
        };
        return (
            <div className="main-content"><div className="content-container"><div className=""><div className="row"><div className="col-md-12 compTitle paddingLeft0">
                <div id="divComments">
                <h2>Comments ({this.state.addedComments.length})
                </h2>

                <div id="comments" className="panel-body">
                    <div className="comment-block comment-post">
                        <div className="form-group">
                            <label>Add comment:</label>
                            <textarea
                                className="form-control"
                                value={this.state.comments}
                                name="comments"
                                onChange={this.onChangeControls}>
                            </textarea>
                            <span
                            className={this.state.errorPost ?
                            "showElem req" : 
                            "hideElem"}>Please provide comments</span>
                        </div>
                        <button
                        type="button"
                        className="btn btn-default post-btn"
                        onClick={()=>this.addCommentReply("Comment#" + this.state.itemID) }>Post Comment</button>

                    </div>
                    {
                        this.state.addedComments.length > 0 ? 
                        <div>
                        {
                            this.state.addedComments.map((file,index) => {
                            return <div>
                                    <div className="comment-block">
                                        <Persona style={{cursor: "pointer" }}
                                            primaryText={file.objParent.Author.Title}
                                            size={PersonaSize.size24}
                                            presence={PersonaPresence.none}
                                            imageUrl={`/_layouts/15/userphoto.aspx?size=S&accountname=${file.objParent.Author.UserName}`}
                                        />
                                        <time
                                            className="posted-date comment-people-dg"
                                            title={new
                                            Date(file.objParent.Created).toLocaleString("en-US",
                                            dateformate)}>
                                            {/*<Moment fromNow>{item.Created.toString()}</Moment>*/}
                                            {new Date(file.objParent.Created).toLocaleString("en-US",
                                            dateformate)}
                                        </time>
                                        <p
                                            className="comment-people-dg"
                                            dangerouslySetInnerHTML={{
                                            __html: file.objParent.CommentBody }}>
                                        </p>
                                        {/*<button
                                            type="button"
                                            className="btn btn-default"
                                            onClick={()=> { this.setDisplayPost(file.objParent.ID)}} id={file.objParent.ID +"_id"}>Reply
                                        </button>*/}
                                        <a onClick={()=> { this.setDisplayPost(file.objParent.ID)}} id={file.objParent.ID +"_id"} href="#" className="reply-btn">
                                            <i className="ms-Icon ms-Icon--Reply x-hidden-focus" title="Reply" aria-hidden="true"></i> Reply
                                        </a>
                                        <div
                                            className={this.state.displayReplyBlock ==
                                            file.objParent.ID ?
                                            "showElem child-txtarea" : 
                                            "hideElem"}>
                                            <textarea
                                                className="form-control"
                                                value={this.state.replyToComments}
                                                name="replyToComments"
                                                onChange={this.onChangeControls}
                                                />
                                            <br></br>
                                            <span
                                                className={this.state.displayReplyBlock ==
                                                file.objParent.ID &&
                                                this.state.errorReply ?
                                                "showElem req" : 
                                                "hideElem"}>Please provide comments</span>
                                            <button
                                                type="button"
                                                className="btn btn-default post-btn"
                                                onClick={()=> { this.addCommentReply("Reply#" + file.objParent.ID); }}
                                                id={file.objParent.ID +"_id"}>Post Reply</button>
                                        </div>
                                    </div>
                                    {
                                        file.objReplies.length > 0 ? 
                                        <div>
                                            {
                                                file.objReplies.map((reply,key) => {
                                                return  <div className="child-replies">
                                                            <div>
                                                                <Persona style={{cursor: "pointer" }}
                                                                    primaryText={reply.Author.Title}
                                                                    size={PersonaSize.size24}
                                                                    presence={PersonaPresence.none}
                                                                    imageUrl={`/_layouts/15/userphoto.aspx?size=S&accountname=${reply.Author.UserName}`}
                                                                />
                                                                <time
                                                                    className="posted-date comment-people-dg"
                                                                    title={new
                                                                    Date(reply.Created).toLocaleString("en-US",
                                                                    dateformate)}>
                                                                    {/*<Moment fromNow>{item.Created.toString()}</Moment>*/}
                                                                    {new Date(reply.Created).toLocaleString("en-US",
                                                                    dateformate)}
                                                                </time>
                                                            </div>
                                                            <div>
                                                            <p
                                                                className="comment-people-dg"
                                                                dangerouslySetInnerHTML={{
                                                                __html: reply.CommentBody }}></p>
                                                        </div>
                                                    </div>
                                                })
                                            }
                                        </div> : null
                                    }
                                </div>
                                }
                            )}
                        </div> : null
                    }
                    </div>
                </div>
            </div></div></div></div></div>
        );
    }
    
}



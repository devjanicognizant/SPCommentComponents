import * as React from 'react';
import styles from './Comments.module.scss';
import { ICommentsProps } from './ICommentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Persona, PersonaSize, PersonaPresence } from "office-ui-fabric-react/lib/Persona";
import Def, { MoveOperations } from "sp-pnp-js";

export default class Comments extends React.Component<ICommentsProps, {
  newComment: string,
    messageContentTypeId: string,
    numberOfItemsToShow: number,
    defaultNumberOfItemsToShow: number

}> {
  private textAreaObject: any;

    constructor(props) {
        super(props);
        this._showOldPost = this._showOldPost.bind(this);

        this.state = {
            newComment: "",
            messageContentTypeId: "",
            numberOfItemsToShow: 10,
            defaultNumberOfItemsToShow: 10
        };

        this._comment = this._comment.bind(this);
        this._postComment = this._postComment.bind(this);
        this._loadMore = this._loadMore.bind(this);
        this._loadMoreSection = this._loadMoreSection.bind(this);
    }

    public componentWillMount() {
        // Def.sp.web.lists.getById(this.props.listName).contentTypes
        //     .filter("Name eq 'Message'")
        //     .get()
        //     .then((respContentType) => {
        //         this.setState({
        //             messageContentTypeId: respContentType[0].Id.StringValue
        //         });
        //     });
    }
    //Load more (3 comments)
    private _loadMore() {
        this.setState({
            numberOfItemsToShow: this.state.numberOfItemsToShow + 3
        });
    }
    //Show all comments
    private _showOldPost() {
        var items = this.props.PostItem.slice(0, this.state.numberOfItemsToShow);
        if (this.props.PostItem.length > 0) {
            return (
                <div >
                    {items.map((post, index) => {
                        return (
                            <div id="allComments" key={index} className="ms-fadeIn500">
                                <div id="UserComments">
                                    <div className="comment-cont" key={"post" + index}>
                                        <div id="commentDesc" dangerouslySetInnerHTML={{ __html: post.Body }}></div>
                                        <div className="row">
                                            <div className="col-sm-6">
                                                <b>{"Replied by "}</b>
                                                <a href={"sip:" + post.Author.SipAddress} style={{ textDecoration: "none" }}>
                                                    <Persona
                                                        style={{ cursor: "pointer" }}
                                                        primaryText={post.Author.Title}
                                                        size={PersonaSize.size24}
                                                        presence={PersonaPresence.none}
                                                        imageUrl={`/_layouts/15/userphoto.aspx?size=S&accountname=${post.Author.UserName}`}
                                                    />
                                                </a>
                                            </div>
                                            <div className="col-sm-6">
                                                <b>{"Replied on "}</b>
                                                {new Date(post.Created).toDateString() + " " + new Date(post.Created).toLocaleTimeString()}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                        );
                    })}
                    {this._loadMoreSection(items)}
                </div>
            );
        }
        else {
            return "";
        }
    }
    //Load More Section
    private _loadMoreSection(items: Object[]) {
        if (items.length < this.props.PostItem.length) {
            return (
                <div>
                    <div className="clearfix"></div>
                    <br />
                    <div className="row">
                        <button className="btn btn-dark" onClick={this._loadMore}>Load More</button>
                    </div>
                </div>
            );
        }
    }
    //Text Area
    private _comment(event) {
        this.textAreaObject = event.target;
        var comment = event.target.value;
        this.setState({
            newComment: comment
        });
    }
    //Add comment
    private _postComment(event) {
        if (this.state.newComment.trim() !== "") {
            let body: string = this.state.newComment.replace(/\n/g, "<br />");
            Def.sp.web.lists.getById(this.props.listName).items.add({
                "Body": body,
                'FileSystemObjectType': 0,
                "ContentTypeId": this.state.messageContentTypeId,
                "ParentItemID": this.props.parentItemId
            })
                .then((resp) => {
                    resp.item.select("FileRef", "FileDirRef")
                        .get()
                        .then((respNewPost) => {

                            var fileUrl = respNewPost.FileRef;
                            var fileDirRef = respNewPost.FileDirRef;

                            Def.sp.web.lists.getById(this.props.listName).items.getById(this.props.parentItemId).folder
                                .get()
                                .then((respParentFolder) => {
                                    var folderUrl = respParentFolder.ServerRelativeUrl;
                                    var moveFileUrl = fileUrl.replace(fileDirRef, folderUrl);
                                    Def.sp.web.getFileByServerRelativeUrl(fileUrl).moveTo(moveFileUrl, MoveOperations.Overwrite)
                                        .then((res) => {
                                            this.textAreaObject.value = "";
                                            this.setState({
                                                newComment: ""
                                            });
                                            this.props.parentObject.getAllPost(this.props.parentObject, this.props.parentItemId);
                                        })
                                        .catch((error) => {
                                            console.error(error);
                                        });
                                });
                        })
                        .catch((error) => {
                            console.error(error);
                        });

                })
                .catch((error) => {
                    console.error("Error occurred while adding reply", error);
                });
        }
        else {
            alert("No Comments found");
        }
    }

  public render(): React.ReactElement<ICommentsProps> {
    /*return (
      <div className={ styles.comments }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );*/

    return (
            <div className="">
                {/*{this.props.HasAddPermission === true ?*/}
                    <div>
                        <div className="section-title">
                            <h1 className="title">Post Your Comments</h1>
                        </div>
                        <div id="commentInput">
                            <textarea name="textComment" id="textComment" cols={50} rows={5} onChange={this._comment}></textarea>
                        </div>
                        <span id="enterCommentSubmit">
                            <span>
                                <button className="btn btn-dark" onClick={this._postComment}>Post Your Comment</button>
                            </span>
                        </span>
                    </div>
                    {/*: ""
                }*/}

                {/*{this._showOldPost()}*/}
            </div>
        );
  }
}

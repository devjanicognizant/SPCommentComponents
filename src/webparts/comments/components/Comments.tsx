import * as React from 'react';
import styles from './Comments.module.scss';
import { ICommentsProps } from './ICommentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Persona, PersonaSize, PersonaPresence } from "office-ui-fabric-react/lib/Persona";
import Def, { MoveOperations } from "sp-pnp-js";
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

export default class Comments extends React.Component<ICommentsProps, {
    newComment: string,
    messageContentTypeId: string,
    numberOfItemsToShow: number,
    defaultNumberOfItemsToShow: number,
    allPosts:any[],
    parentItemId:number
}> {
  private textAreaObject: any;

    constructor(props) {
        super(props);
        this._showOldPost = this._showOldPost.bind(this);

        this.state = {
            newComment: "",
            messageContentTypeId: "",
            numberOfItemsToShow: 10,
            defaultNumberOfItemsToShow: 10,
            allPosts:[],
            parentItemId:0
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
        var queryParameters = new UrlQueryParameterCollection(window.location.href);
        var parentListFieldName = this.props.parentItemIdFieldName;
        var itemId = queryParameters.getValue("ComponentID");
        //itemId="6";
        if (itemId) {
        let id = parseInt(itemId);
        this.setState({
            parentItemId : id
            });

            this.getAllPost(id);
        }
        
    }
    //Load more (3 comments)
    private _loadMore(event) {
        this.setState({
            numberOfItemsToShow: this.state.numberOfItemsToShow + 3
        });
        event.preventDefault();
    }
    //Get All Discussions
    public getAllPost(itemId:any) {
        var parentObject =  this;
        var commentListName = this.props.listName;
        var parentListFieldName = parentObject.props.parentItemIdFieldName;
        Def.sp.web.lists.getByTitle(parentObject.props.listName).items
        .select("*", "Author/UserName", "Author/SipAddress", "Author/Title", "Author/Id", "Author/EMail")
        .expand("Author")
        .filter(parentListFieldName+" eq '" + itemId + "'")
        .orderBy("Created", false)
        .getAll(4000)
        .then((repliedPost) => {
            repliedPost = this.sortByDate(repliedPost, "Created", false);
            parentObject.setState({
            allPosts : repliedPost
            });
        });
    }
    //Sort by Date
    private sortByDate(arrayObject: any[], key, ascending: boolean = true) {
        arrayObject.sort((a, b) => {
        var aDate = new Date(a[key]);
        var bDate = new Date(b[key]);
        if (aDate > bDate) {
            return 1;
        }
        if (aDate < bDate) {
            return -1;
        }
        return 0;
        });
        if (!ascending) {
        arrayObject = arrayObject.reverse();
        }
        return arrayObject;
    }


    //Show all comments
    private _showOldPost() {
        var items = this.state.allPosts.slice(0, this.state.numberOfItemsToShow);
        if (this.state.allPosts.length > 0) {
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
        if (items.length < this.state.allPosts.length) {
            return (
                <div>
                    <div className="clearfix"></div>
                    <br />
                    <div className="row">
                        <button className="btn btn-dark" onClick={(e) => this._loadMore(e)}>Load More</button>
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
            var parentListFieldName = this.props.parentItemIdFieldName;
            Def.sp.web.lists.getByTitle(this.props.listName).items.add({
                "Body": body,
                "ParentItemId": this.state.parentItemId
            })
            .then((resp) => {
                this.textAreaObject.value = "";
                this.setState({
                    newComment: ""
                });
                this.getAllPost(this.state.parentItemId);
            })
            .catch((error) => {
                console.error("Error occurred while adding reply", error);
            });
        }
        else {
            alert("No Comments found");
        }
        event.preventDefault();
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
                                <button className="btn btn-dark" onClick={(e) => this._postComment(e)}>Post Your Comment</button>
                            </span>
                        </span>
                    </div>
                    {/*: ""
                }*/}

                {this._showOldPost()}
            </div>
        );
  }
}

import * as React from 'react';
import styles from './RecommendedArticles.module.scss';
import { IRecommendedArticlesProps } from './IRecommendedArticlesProps';
import { IRecommendedArticlesState } from './IRecommendedArticlesState';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all"; 
import { Web } from '@pnp/sp/webs';

const UserInterestListName = "UserInterestList";
const TestListName = "Test";
export default class RecommendedArticles extends React.Component<IRecommendedArticlesProps, IRecommendedArticlesState> {
  constructor(props: IRecommendedArticlesProps){
    super(props);
    this.state = {
      allPages: []
    }
  }
  
  public render(): React.ReactElement<IRecommendedArticlesProps> {
    return (
      <div className={ styles.recommendedArticles }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              {/* <span className={ styles.title }>Welcome to SharePoint!</span> */}
              {/* <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a> */}
              {/* <div className={styles.wrapper}> */}
                <div className={styles.wrapper}>
                    {this.state.allPages.map(item => {
                      return <div className={`styles.one styles.container`}>{item.Title}</div>
                      })}
                  </div>  
              {/* </div> */}
            </div>
          </div>
        </div>
      </div>
    );
  }

  componentDidMount() {
    const RootWeb = Web("https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/");

    this.getUserObject().then(items => {
      this.getUserInterests(items).then(interests => {
        this.getAllPages(interests).then(pages => {
          //this.addItem();
           this.setState({
             allPages:pages
           }); 
        })
      });
    });
  }

  //Get User Email
  private getUserObject = async() =>{
    try{
      const userObject: any = await sp.web.currentUser.get();
      console.log(userObject.Email);
      return userObject.Email;
    }
    catch(e){
      console.log(e);
    }
  }

  //Add an item in Test List
  private addItem = async() =>{
    try{
      console.log("Adding an item in test list");
      const itemAdded: any = await sp.web.lists.getByTitle(TestListName).items.add({
          Title: "Nishikant" 
      });
        console.log("New item addition completed");
        return itemAdded;
      }
      catch(e){
        console.error(e);
      }
  }
  
  //Get User Interest based on logged in user
  private getUserInterests = async (email) => {
    try{
          console.log("Getting user interests");
          const userInterests: any = await sp.web.lists.getByTitle(UserInterestListName).items
            .select("User_x0020_Interests")
            .filter(`Title eq '${email}'`)
            .get();
            console.log("Gettin user interersts",userInterests);
            return userInterests;
    }
    catch(e){
      console.error(e);
    }
  }

  //Get all pages
  private getAllPages = async (interests) => {
    try {

      let interestFilterURL = "";
      let interestFilterURLUpdated = "";
      interests[0].User_x0020_Interests.map(item => {
        interestFilterURL += `(User_x0020_Interests eq '${item}') or`
        interestFilterURLUpdated = interestFilterURL.slice(0,-3);
        console.log(interestFilterURLUpdated);
      })
      const items: any[] = await sp.web.lists.getByTitle("Site Pages").items
      .select("Title")
      //.filter("User_x0020_Interests eq Sports")
      .filter(interestFilterURLUpdated)
      .get();
      console.log(items);
      return items;
    }
    catch (e) {
      console.error(e);
    }
  }
}

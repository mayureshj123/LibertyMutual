import * as React from 'react';
import styles from './LibertyPages.module.scss';
import { ILibertyPagesState } from './ILibertyPagesState';
import { ILibertyPagesProps } from './ILibertyPagesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web } from '@pnp/sp/webs';

export default class LibertyPages extends React.Component<ILibertyPagesProps, ILibertyPagesState> {

  constructor(props: ILibertyPagesProps) {
    super(props);
    this.state = {
      rootPages: [],
      subSitePages: []
    }
  }
  public render(): React.ReactElement<ILibertyPagesProps> {
    const allPages = this.state.rootPages.concat(this.state.subSitePages);
    return (
      <div className={styles.libertyPages}>
        <div className={`container`}>
          <ul className="nav nav-tabs">
            <li className="active"><a data-toggle="tab" href="#menu1">{`All (${this.state.subSitePages.length + this.state.subSitePages.length})`}</a></li>
            <li><a data-toggle="tab" href="#menu2">{`On Demand (${this.state.rootPages.length})`}</a></li>
            <li><a data-toggle="tab" href="#menu2">{`Surety You (${this.state.subSitePages.length})`}</a></li>
          </ul>

          <div className="tab-content">
            <div id="menu1" className="tab-pane fade in active">
              {allPages.map(item => {
                return <p>
                  {item.Title}
                </p>
              })}
            </div>
            <div id="menu2" className="tab-pane fade">
              {this.state.rootPages.map(item => {
                return <p>
                  {item.Title}
                </p>
              })}          </div>
            <div id="menu3" className="tab-pane fade">
              {this.state.subSitePages.map(item => {
                return <p>
                  {item.Title}
                </p>
              })}
            </div>
          </div>
        </div>
      </div>
    );
  }
  componentDidMount() {
    const RootWeb = Web("https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/");
    const ChildWeb = Web("https://sirisrdalabs.sharepoint.com/sites/LibertyMutual/");

    this.getAllItems(RootWeb).then(items => {
      this.setState({
        rootPages: items
      });
    });

    this.getAllItems(ChildWeb).then(items => {
      this.setState({
        subSitePages: items
      });
    });

  }

  //Get all items
  private getAllItems = async (rerewr) => {
    try {
      const items: any[] = await rerewr.lists.getByTitle("Site Pages").items.get();
      return items;
    }
    catch (e) {
      console.error(e);
    }
  }
}

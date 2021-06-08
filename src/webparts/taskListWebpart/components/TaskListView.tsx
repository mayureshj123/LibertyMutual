import * as React from 'react';
import styles from './TaskListWebpart.module.scss';
import { ITaskListViewProps } from './ITaskListViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";


export interface ITaskListViewState {
  items: any[];
  lists: string | string[];
  loggedInUserID: string;
  description : string;
}
export default class TaskListView extends React.Component<ITaskListViewProps, ITaskListViewState> {
  constructor(props: ITaskListViewProps, state: ITaskListViewState) {
    super(props);
    this.state = {
      items: [],
      // lists: this.props.lists,
      lists: "7209dab0-fbc4-4c73-9533-2816f6727ff6",
      loggedInUserID: "",
      description : this.props.description

    }
  }
  public render(): React.ReactElement<ITaskListViewProps> {
    return (
      <React.Fragment>
        {/* <div className= {styles.taskListView}> */}
        <div>
        {/* <h2 onMouseOver={this.MouseOver} onMouseOut={this.MouseOut} >My Task</h2> */}
        <h2>My Task</h2>
          <table>
            {/* <tr><th>ID</th></tr> */}
            {this.state.items.map((item: any, index: number) => (
              <tr>
                <tr>
                  <td>
                    {/* <img src="img_girl.jpg" alt="Avatar" className={styles.avatar} /> */}
                    <div className={styles.taskCircle}>{(item.Title).toString().slice(0, 2).toUpperCase()}</div>

                  </td>
                  <td>
                    <strong>{item.Title}</strong>
                  </td>
                </tr>
                <tr>
                  <td></td>
                  <td><span className={styles.taskContent} >{item.Body}</span>
                  </td>
                </tr>
                <tr>
                  <td></td>
                  {/* <td><span className={styles.taskLastModified}>{(item.CreatedSince).toString().split('.')[0] >= 1 ?  "Today" : item.CreatedSince + " Days Ago"}</span></td> */}

                  <td><span className={styles.taskLastModified}>{(item.Created).toString().slice(8, 10) + '-' + (item.Created).toString().slice(5, 7) + '-' + (item.Created).toString().slice(0, 4)}</span></td>

                </tr>
              </tr>
            ))
            }
          </table>
        </div>
      </React.Fragment>
    );
  }


  public componentDidMount() {
    sp.web.currentUser.get().then((user) => {
      console.log(user.Id);
      this.setState({
        loggedInUserID: user.Id.toString(),
        lists: this.state.lists
      });
      this.getItems(this.state.loggedInUserID);

    });
    // console.log(this.props.context );
  }

  private async getItems(loggedInUserID) {
    const items = await sp.web.lists.getById(this.state.lists.toString()).items.select()
      .expand()
      .filter(" AssignedTo eq " + loggedInUserID + " ")
      .top(5)
      .orderBy("Created", false)
      .get();
    this.setState({
      items: items ? items : []
    });
    console.log(this.state.items);
    //console.log(this.state);
    //console.log(this.props.lists);
    //console.log(this.props.context);

  }
  // public   MouseOver(event) {
//   event.target.style.visible = 'red';
//   event.target.style.display = "block";
// }
// public MouseOut(event){
//   event.target.style.background="";
//   event.target.style.display = "none";

// }
}

import * as React from 'react';
import styles from './TaskListWebpart.module.scss';
import { ITaskListWebpartProps } from './ITaskListWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import TaskListView from './TaskListView';

export default class TaskListWebpart extends React.Component<ITaskListWebpartProps, {}> {
  public render(): React.ReactElement<ITaskListWebpartProps> {
    return (
      <div className={ styles.taskListWebpart }>
        {/* <div className={ styles.container }>
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
        </div> */}
        <TaskListView description = {this.props.description} context = {this.props.context}></TaskListView>
      </div>
    );
  }
}

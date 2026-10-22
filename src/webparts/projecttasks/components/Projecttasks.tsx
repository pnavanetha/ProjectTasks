import * as React from 'react';
import styles from './Projecttasks.module.scss';
import type { IProjecttasksProps } from './IProjecttasksProps';
// import { escape } from '@microsoft/sp-lodash-subset';
// import EmployeeForm from './Forms/Employee';
// import HolidayForm from './Forms/Holidays';
import TasksBacklogs from './Forms/TasksBacklogs';

export default class Projecttasks extends React.Component<IProjecttasksProps> {
  public render(): React.ReactElement<IProjecttasksProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (
      <section className={`${styles.projecttasks} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        {/* <h2>Employee Form</h2>
        <EmployeeForm /> */}
        {/* <h2>Holiday Form</h2>
        <HolidayForm /> */}
        <h2>Task Backlogs</h2>
        <TasksBacklogs />

      </section>
    );
  }
}

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from '@microsoft/teams-js';

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      context: {},
    };
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount() {
    // Get the user context from Teams and set it in the state
    microsoftTeams.getContext((context, error) => {
      this.setState({
        context: context,
      });
    });
    microsoftTeams.registerOnThemeChangeHandler((theme) => {
      if (theme !== this.state.context.theme) {
        console.log('Not match!');
        this.setState({ theme });
      }
    });
    // Next steps: Error handling using the error object
  }

  render() {
    const isTheme = this.state.theme === 'default';
    console.log('theme: ', this.state.context);
    console.log('isTheme: ', isTheme);

    let userName =
      Object.keys(this.state.context).length > 0
        ? this.state.context['upn']
        : '';
    console.log('username: ', userName);

    return (
      <div className={isTheme ? 'Lite' : 'Dark'}>
        <h1>Important Contacts</h1>
        <ul>
          <li>
            Help Desk:{' '}
            <a href="mailto:support@company.com">support@company.com</a>
          </li>
          <li>
            Human Resources: <a href="mailto:hr@company.com">hr@company.com</a>
          </li>
          <li>
            Facilities:{' '}
            <a href="mailto:facilities@company.com">facilities@company.com</a>
          </li>
        </ul>
      </div>
    );
  }
}
export default Tab;

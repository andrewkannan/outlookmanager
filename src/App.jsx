import React from 'react';
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { Login } from './components/Login';
import { KanbanBoard } from './components/KanbanBoard';

function App() {
  return (
    <div className="app-container">
      <AuthenticatedTemplate>
        <KanbanBoard />
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <Login />
      </UnauthenticatedTemplate>
    </div>
  );
}

export default App;

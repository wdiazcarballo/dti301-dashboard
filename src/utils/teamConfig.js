// teamConfig.js
import * as microsoftTeams from "@microsoft/teams-js";
import { PublicClientApplication } from "@azure/msal-browser";

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: process.env.NEXT_PUBLIC_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.NEXT_PUBLIC_TENANT_ID}`,
    redirectUri: process.env.NEXT_PUBLIC_REDIRECT_URI,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: true
  }
};

// Initialize Teams SDK
export const initializeTeamsContext = async () => {
  try {
    await microsoftTeams.app.initialize();
    const context = await microsoftTeams.app.getContext();
    return context;
  } catch (error) {
    console.error("Error initializing Teams:", error);
    return null;
  }
};

// Initialize MSAL
export const msalInstance = new PublicClientApplication(msalConfig);

// Updated Graph API scopes for education
export const graphScopes = [
  "User.Read",
  "Team.ReadBasic.All",
  "ChannelMessage.Read.All",
  "EduAssignments.Read",
  "EduAssignments.ReadWrite",
  "EduAssignments.ReadBasic",
  "EduAssignments.ReadWriteBasic"
];

// Get class assignments
export const getClassAssignments = async (accessToken, classId) => {
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    return await response.json();
  } catch (error) {
    console.error("Error fetching assignments:", error);
    return null;
  }
};

// Get student submissions for an assignment
export const getStudentSubmissions = async (accessToken, classId, assignmentId) => {
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/submissions`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    return await response.json();
  } catch (error) {
    console.error("Error fetching submissions:", error);
    return null;
  }
};

// Get student profile with educational info
export const getStudentProfile = async (accessToken) => {
  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/education/me",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    return await response.json();
  } catch (error) {
    console.error("Error fetching profile:", error);
    return null;
  }
};

// Submit mindfulness reflection
export const submitReflection = async (accessToken, classId, assignmentId, reflection) => {
  try {
    // Create submission resource
    const submissionResponse = await fetch(
      `https://graph.microsoft.com/v1.0/education/classes/${classId}/assignments/${assignmentId}/submissions`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          resourcesFolderUrl: null,
          submittedDateTime: new Date().toISOString(),
          content: {
            text: reflection,
            contentType: "text"
          }
        }),
      }
    );
    
    if (!submissionResponse.ok) {
      throw new Error('Failed to submit reflection');
    }

    return await submissionResponse.json();
  } catch (error) {
    console.error("Error submitting reflection:", error);
    return null;
  }
};

// Get class details
export const getClassDetails = async (accessToken, classId) => {
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/education/classes/${classId}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    return await response.json();
  } catch (error) {
    console.error("Error fetching class details:", error);
    return null;
  }
};
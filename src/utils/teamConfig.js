import * as microsoftTeams from "@microsoft/teams-js";
import { PublicClientApplication } from "@azure/msal-browser";

// MSAL configuration - You'll need to update these values
const msalConfig = {
  auth: {
    clientId: process.env.NEXT_PUBLIC_CLIENT_ID, // From Azure AD app registration
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

// Graph API scopes needed
export const graphScopes = [
  "User.Read",
  "Team.ReadBasic.All",
  "ChannelMessage.Read.All",
  "Assignment.Read.All"
];

// Get Teams assignments
export const getTeamsAssignments = async (accessToken, teamId) => {
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/assignments`,
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

// Get student profile
export const getStudentProfile = async (accessToken) => {
  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me",
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

// Save mindfulness data to Teams channel
export const saveMindfulnessData = async (accessToken, teamId, channelId, data) => {
  try {
    const cardData = {
      contentType: "application/vnd.microsoft.card.adaptive",
      content: {
        type: "AdaptiveCard",
        version: "1.4",
        body: [
          {
            type: "TextBlock",
            text: "Mindfulness Progress Update",
            weight: "bolder",
            size: "medium"
          },
          {
            type: "FactSet",
            facts: [
              {
                title: "Meditation Minutes:",
                value: data.meditationMinutes.toString()
              },
              {
                title: "Reflection Complete:",
                value: data.reflectionComplete ? "Yes" : "No"
              }
            ]
          }
        ]
      }
    };

    await fetch(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          body: {
            contentType: "html",
            content: "Mindfulness Progress Update"
          },
          attachments: [cardData]
        }),
      }
    );
  } catch (error) {
    console.error("Error saving mindfulness data:", error);
  }
};
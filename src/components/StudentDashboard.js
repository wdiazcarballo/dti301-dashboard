import React, { useState, useEffect } from 'react';
import { Card, CardHeader, CardContent } from '@/components/ui/card';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { Heart, Brain, Dumbbell } from 'lucide-react';
import { initializeTeamsContext, msalInstance, graphScopes, getTeamsAssignments, getStudentProfile } from './teamConfig';

const StudentDashboard = () => {
  const [teamsContext, setTeamsContext] = useState(null);
  const [assignments, setAssignments] = useState([]);
  const [userProfile, setUserProfile] = useState(null);
  const [language, setLanguage] = useState('th');

  useEffect(() => {
    const initializeTeams = async () => {
      // Initialize Teams context
      const context = await initializeTeamsContext();
      setTeamsContext(context);

      if (context) {
        try {
          // Get access token
          const account = msalInstance.getAllAccounts()[0];
          const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: graphScopes,
            account: account
          });

          // Get assignments and profile
          const assignmentsData = await getTeamsAssignments(
            tokenResponse.accessToken,
            context.team.groupId
          );
          const profileData = await getStudentProfile(tokenResponse.accessToken);

          setAssignments(assignmentsData);
          setUserProfile(profileData);
        } catch (error) {
          console.error("Error initializing dashboard:", error);
        }
      }
    };

    initializeTeams();
  }, []);

  // Render loading state if not initialized
  if (!teamsContext) {
    return <div className="p-6">Loading Teams context...</div>;
  }

  return (
    <div className="p-6 bg-gray-50 min-h-screen">
      <div className="max-w-7xl mx-auto">
        {/* Student Info */}
        <Card className="mb-8">
          <CardContent className="p-6">
            <div className="flex items-center space-x-4">
              <div>
                <h2 className="text-xl font-bold">
                  {userProfile?.displayName || 'Loading...'}
                </h2>
                <p className="text-gray-500">DTI 301 - Professional Ethics</p>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Mindful Self Discipline Tracking (Teams doesn't have this) */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
          <Card>
            <CardContent className="p-6">
              <div className="flex items-center space-x-4">
                <Heart className="h-8 w-8 text-red-500" />
                <div>
                  <p className="text-sm text-gray-500">
                    {language === 'en' ? 'Heart (Love)' : 'หัวใจ (ความรัก)'}
                  </p>
                  <p className="text-2xl font-bold">8/10</p>
                </div>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-6">
              <div className="flex items-center space-x-4">
                <Dumbbell className="h-8 w-8 text-purple-500" />
                <div>
                  <p className="text-sm text-gray-500">
                    {language === 'en' ? 'Power (Body)' : 'พลัง (ร่างกาย)'}
                  </p>
                  <p className="text-2xl font-bold">7/10</p>
                </div>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-6">
              <div className="flex items-center space-x-4">
                <Brain className="h-8 w-8 text-blue-500" />
                <div>
                  <p className="text-sm text-gray-500">
                    {language === 'en' ? 'Wisdom (Mind)' : 'ปัญญา (จิตใจ)'}
                  </p>
                  <p className="text-2xl font-bold">9/10</p>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Weekly Reflection (Integrated with Teams Assignments) */}
        <Card className="mb-8">
          <CardHeader>
            <h2 className="text-xl font-semibold">
              {language === 'en' ? 'Weekly Reflection' : 'การสะท้อนรายสัปดาห์'}
            </h2>
          </CardHeader>
          <CardContent>
            <textarea 
              className="w-full h-32 p-4 border rounded-lg resize-none"
              placeholder={language === 'en' ? 
                "Write your weekly reflection here..." : 
                "เขียนการสะท้อนประจำสัปดาห์ของคุณที่นี่..."}
              onChange={(e) => {
                // Save to Teams channel when appropriate
              }}
            />
          </CardContent>
        </Card>

        {/* Progress Visualization */}
        <Card className="mb-8">
          <CardHeader>
            <h2 className="text-xl font-semibold">
              {language === 'en' ? 'Course Progress' : 'ความก้าวหน้าในรายวิชา'}
            </h2>
          </CardHeader>
          <CardContent>
            <div className="h-96">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={assignments}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="dueDateTime" />
                  <YAxis />
                  <Tooltip />
                  <Legend />
                  <Line 
                    type="monotone" 
                    dataKey="completionRate" 
                    stroke="#3b82f6" 
                    name={language === 'en' ? 'Completion' : 'ความสำเร็จ'} 
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default StudentDashboard;
import React, { useEffect } from "react";
import { app } from "@microsoft/teams-js";

function App() {
  useEffect(() => {
    const getUser = async () => {
      try {
        await app.initialize();

        const context = await app.getContext();

        window.alert("User Details:", context.user);

        /*
        {
            id: "aad-object-id",
            displayName: "Amrut Telavane",
            userPrincipalName: "user@company.com"
        }
        */
      } catch (error) {
        console.error("Error getting Teams context:", error);
      }
    };

    getUser();
  }, []);

  return (
    <div>
      <h1>Microsoft Teams React App</h1>
    </div>
  );
}

export default App;

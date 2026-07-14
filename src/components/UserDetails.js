import { app } from "@microsoft/teams-js";

async function getUser() {
  await app.initialize();

  const context = await app.getContext();

  console.log(context.user);

  /*
  {
      id: "aad-object-id",
      displayName: "Amrut Telavane",
      userPrincipalName: "user@company.com"
  }
  */
}

getUser();
export default getUser()
import { sp } from '@pnp/sp/presets/all';

export const checkUserRole = async () => {
  const currentUser = await sp.web.currentUser.get();
  const userEmail = currentUser.Email;

  const userItems = await sp.web.lists.getByTitle("UsersList")
    .items
    .filter(`User/EMail eq '${userEmail}' and Actif eq 1`)
    .expand("User")
    .select("Role", "User/Title")
    .top(1)
    .get();

  if (userItems.length === 0) {
    return { isAllowed: false, role: null };
  }

  return {
    isAllowed: true,
    role: userItems[0].Role
  };
};

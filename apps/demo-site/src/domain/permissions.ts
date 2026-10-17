/**
 * Authorization seam. The POC lets every signed-in user edit every file; the
 * real app answers from its ACLs. When a permission is removed there, the app
 * must also revoke the user's DAV sessions for that file
 * (`DavPolicy.revokeSession`), because revoked always means 401.
 */
export interface Permissions {
  canWrite(userId: string, fileId: string): boolean;
}

export const everyoneCanWrite: Permissions = {
  canWrite: () => true,
};

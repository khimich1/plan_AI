export type UserRole = "admin" | "manager" | "production" | string;

export type CurrentUser = {
  id: number;
  username: string;
  role: UserRole;
  manager_id?: number | null;
  is_active?: number | boolean;
  created_at?: string;
};

export type CurrentUserResponse = {
  user: CurrentUser;
};

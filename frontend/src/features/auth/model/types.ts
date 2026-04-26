export type AuthUser = {
  id: number;
  username: string;
  role: string;
  manager_id: number | null;
  is_active: boolean | number;
  created_at?: string;
};

export type LoginPayload = {
  username: string;
  password: string;
};

export type MeResponse = {
  user: AuthUser;
};

export type LoginResponse = {
  user: AuthUser;
};

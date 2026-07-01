import React, { createContext, useContext, useState } from "react";

// Shared scopes used by the app
export const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets.readonly"
];

interface AuthContextType {
  user: any | null;
  token: string | null;
  isLoggingIn: boolean;
  login: () => Promise<void>;
  logout: () => Promise<void>;
  initialized: boolean;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<any | null>({ displayName: "User" });
  const [token, setToken] = useState<string | null>("dummy-token");
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [initialized, setInitialized] = useState(true);

  const login = async () => {};
  const logout = async () => {};

  return (
    <AuthContext.Provider value={{ user, token, isLoggingIn, login, logout, initialized }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
};

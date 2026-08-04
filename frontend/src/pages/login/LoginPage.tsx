import { useState } from "react";
import { Navigate, useNavigate } from "react-router";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { getErrorMessage } from "@/shared/lib/apiError";
import { defaultRouteForRole } from "@/shared/lib/roleRoutes";
import { Spinner } from "@/shared/ui/Spinner";

const loginSchema = z.object({
  username: z.string().trim().min(1, "Введите логин"),
  password: z.string().min(1, "Введите пароль"),
});

type LoginFormValues = z.infer<typeof loginSchema>;

export const LoginPage = () => {
  const { user, isLoading, login, isLoggingIn } = useAuth();
  const navigate = useNavigate();
  const [submitError, setSubmitError] = useState<string | null>(null);

  const form = useForm<LoginFormValues>({
    resolver: zodResolver(loginSchema),
    defaultValues: { username: "", password: "" },
  });

  if (isLoading) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
        }}
      >
        <Spinner />
      </div>
    );
  }

  if (user) {
    return <Navigate to={defaultRouteForRole(user.role)} replace />;
  }

  const onSubmit = async (values: LoginFormValues) => {
    setSubmitError(null);
    try {
      const loggedInUser = await login(values);
      navigate(defaultRouteForRole(loggedInUser.role), { replace: true });
    } catch (error) {
      setSubmitError(getErrorMessage(error));
    }
  };

  return (
    <main
      style={{
        minHeight: "100vh",
        display: "grid",
        placeItems: "center",
        padding: "2rem 1rem",
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: 420,
          background: "#ffffff",
          borderRadius: 16,
          border: "1px solid #e4e7ec",
          boxShadow: "0 20px 45px rgba(15, 23, 42, 0.08)",
          padding: "2rem",
        }}
      >
        <div style={{ marginBottom: "1.5rem" }}>
          <div
            style={{
              display: "grid",
              placeItems: "center",
              width: 48,
              height: 48,
              borderRadius: 14,
              background: "linear-gradient(135deg, #2b5cff, #6e7dff)",
              color: "#ffffff",
              fontWeight: 700,
              fontSize: "1.2rem",
              marginBottom: "1rem",
            }}
          >
            S
          </div>
          <h1 style={{ margin: 0, fontSize: "1.5rem" }}>Вход в кабинет</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#667085" }}>
            Введите логин и пароль, чтобы продолжить работу.
          </p>
        </div>

        <form
          onSubmit={form.handleSubmit(onSubmit)}
          style={{ display: "grid", gap: "1rem" }}
          noValidate
        >
          <FieldWrapper
            label="Логин"
            error={form.formState.errors.username?.message}
          >
            <Input
              autoFocus
              autoComplete="username"
              placeholder="Например, admin"
              disabled={isLoggingIn}
              {...form.register("username")}
            />
          </FieldWrapper>

          <FieldWrapper
            label="Пароль"
            error={form.formState.errors.password?.message}
          >
            <Input
              type="password"
              autoComplete="current-password"
              placeholder="Ваш пароль"
              disabled={isLoggingIn}
              {...form.register("password")}
            />
          </FieldWrapper>

          {submitError && <Alert tone="error">{submitError}</Alert>}

          <Button type="submit" fullWidth disabled={isLoggingIn}>
            {isLoggingIn ? "Входим..." : "Войти"}
          </Button>
        </form>
      </div>
    </main>
  );
};

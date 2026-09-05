"use client";

import { Suspense, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";

export default function Login() {
  return (
    <Suspense>
      <FormularioLogin />
    </Suspense>
  );
}

function FormularioLogin() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [enviando, setEnviando] = useState(false);

  async function enviar(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setEnviando(true);
    try {
      const res = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password }),
      });
      if (!res.ok) {
        const json = await res.json().catch(() => ({}));
        throw new Error(json.error || "No se pudo iniciar sesión.");
      }
      router.push(searchParams.get("volver") || "/");
      router.refresh();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setEnviando(false);
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-6" style={{ background: "var(--bg)" }}>
      <div className="card p-10 max-w-sm w-full flex flex-col items-center gap-6 text-center">
        <div>
          <h1 className="text-xl font-semibold">Dashboard CAPEX</h1>
          <p className="text-sm mt-1" style={{ color: "var(--texto-suave)" }}>
            Ingresa la contraseña para continuar.
          </p>
        </div>
        <form onSubmit={enviar} className="w-full flex flex-col gap-3">
          <input
            type="password"
            className="campo"
            placeholder="Contraseña"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoFocus
            disabled={enviando}
          />
          {error && (
            <p className="text-sm text-left" style={{ color: "var(--peligro)" }}>
              {error}
            </p>
          )}
          <button type="submit" className="boton-primario w-full" disabled={enviando || !password}>
            {enviando ? "Entrando…" : "Entrar"}
          </button>
        </form>
      </div>
    </div>
  );
}

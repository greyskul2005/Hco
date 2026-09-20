import type { ReactNode } from "react";
import { cn } from "@/lib/utils";

export function Badge({
  className,
  tone = "navy",
  children,
}: {
  className?: string;
  tone?: "navy" | "blue" | "cyan" | "ok" | "warn" | "err" | "muted" | "inc" | "sus";
  children: ReactNode;
}) {
  return (
    <span
      className={cn(
        "inline-flex items-center rounded-md px-2 py-0.5 text-[11px] font-medium tracking-wide",
        tone === "navy" && "bg-primary text-primary-foreground",
        tone === "blue" && "bg-blue-soft text-primary",
        tone === "cyan" && "bg-cyan-soft text-cyan",
        tone === "ok" && "bg-ok text-ok-fg",
        tone === "warn" && "bg-warn text-warn-fg",
        tone === "err" && "bg-err text-err-fg",
        tone === "muted" && "bg-muted text-muted-foreground",
        tone === "inc" && "bg-inc text-primary",
        tone === "sus" && "bg-sus text-cyan",
        className,
      )}
    >
      {children}
    </span>
  );
}

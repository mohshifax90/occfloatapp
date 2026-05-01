"use client"

import * as React from "react"
import { useTheme } from "next-themes"

function ThemeToggle({ className }: { className?: string }) {
  const { resolvedTheme, setTheme } = useTheme()
  const [mounted, setMounted] = React.useState(false)

  React.useEffect(() => {
    setMounted(true)
  }, [])

  const isDark = mounted && resolvedTheme === "dark"

  return (
    <button
      type="button"
      onClick={() => setTheme(isDark ? "light" : "dark")}
      className={
        "btn btn-sm btn-outline-secondary d-inline-flex align-items-center gap-2" +
        (className ? " " + className : "")
      }
      aria-label="Toggle theme"
      title="Toggle theme (D)"
    >
      <span aria-hidden="true">{isDark ? "\u2600" : "\u263D"}</span>
      <span className="d-none d-sm-inline">{isDark ? "Light" : "Dark"}</span>
    </button>
  )
}

export { ThemeToggle }

import * as React from "react"

export type ButtonVariant =
  | "default"
  | "outline"
  | "secondary"
  | "ghost"
  | "destructive"
  | "link"

export type ButtonSize = "default" | "xs" | "sm" | "lg" | "icon" | "icon-xs" | "icon-sm" | "icon-lg"

type ButtonProps = React.ComponentProps<"button"> & {
  variant?: ButtonVariant
  size?: ButtonSize
}

const variantClassMap: Record<ButtonVariant, string> = {
  default: "btn-primary",
  outline: "btn-outline-secondary",
  secondary: "btn-secondary",
  ghost: "btn-light",
  destructive: "btn-danger",
  link: "btn-link",
}

const sizeClassMap: Record<ButtonSize, string> = {
  default: "",
  xs: "btn-sm",
  sm: "btn-sm",
  lg: "btn-lg",
  icon: "btn-sm",
  "icon-xs": "btn-sm",
  "icon-sm": "btn-sm",
  "icon-lg": "btn-lg",
}

function joinClasses(...parts: Array<string | null | undefined | false>) {
  return parts.filter(Boolean).join(" ")
}

function Button({
  className,
  variant = "default",
  size = "default",
  type = "button",
  ...props
}: ButtonProps) {
  return (
    <button
      type={type}
      className={joinClasses("btn", variantClassMap[variant], sizeClassMap[size], className)}
      {...props}
    />
  )
}

export { Button }

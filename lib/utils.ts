type ClassInput = string | number | false | null | undefined | ClassInput[]

export function cn(...inputs: ClassInput[]): string {
  const parts: string[] = []
  const walk = (value: ClassInput) => {
    if (!value && value !== 0) return
    if (Array.isArray(value)) {
      value.forEach(walk)
      return
    }
    parts.push(String(value))
  }
  inputs.forEach(walk)
  return parts.join(" ")
}

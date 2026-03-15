const IDENTIFIER_PATTERN = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

export const normalizeNameKey = (value: string) => value.trim().toUpperCase();

export const applyCase = (
  value: string,
  mode: "none" | "camelCase" | "snake_case" | "SCREAMING_SNAKE_CASE"
): string => {
  if (mode === "none") return value;
  const tokens = value
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);
  if (tokens.length === 0) return value;
  if (mode === "camelCase") {
    return tokens
      .map((part, idx) =>
        idx === 0
          ? `${part.charAt(0).toLowerCase()}${part.slice(1).toLowerCase()}`
          : `${part.charAt(0).toUpperCase()}${part.slice(1).toLowerCase()}`
      )
      .join("");
  }
  if (mode === "snake_case") return tokens.map((t) => t.toLowerCase()).join("_");
  return tokens.map((t) => t.toUpperCase()).join("_");
};

export const buildUniqueScopedName = (
  baseName: string,
  occupied: Set<string>,
  fallbackName: string
): string => {
  const trimmedBase = baseName.trim() || fallbackName;
  let candidate = trimmedBase;
  let suffix = 2;
  while (occupied.has(normalizeNameKey(candidate))) {
    candidate = `${trimmedBase}_${suffix.toString()}`;
    suffix += 1;
  }
  return candidate;
};

export const extractLambdaArgsFromFormula = (formulaInput: string): string[] => {
  const normalized = formulaInput.trim().replace(/^=/, "").trim();
  if (!/^LAMBDA\s*\(/i.test(normalized)) {
    return [];
  }
  const start = normalized.indexOf("(");
  const end = normalized.lastIndexOf(")");
  if (start < 0 || end <= start + 1) {
    return [];
  }
  const inner = normalized.slice(start + 1, end);
  const parts: string[] = [];
  let depth = 0;
  let inString = false;
  let current = "";
  for (let index = 0; index < inner.length; index += 1) {
    const ch = inner[index];
    if (ch === '"') {
      inString = !inString;
      current += ch;
      continue;
    }
    if (!inString) {
      if (ch === "(") {
        depth += 1;
      } else if (ch === ")") {
        depth = Math.max(0, depth - 1);
      } else if (ch === "," && depth === 0) {
        parts.push(current.trim());
        current = "";
        continue;
      }
    }
    current += ch;
  }
  if (current.trim()) {
    parts.push(current.trim());
  }
  if (parts.length <= 1) {
    return [];
  }
  return parts.slice(0, -1).filter((item) => IDENTIFIER_PATTERN.test(item));
};

export const parseNamedListFormula = (formulaInput: string): string[] => {
  const normalized = formulaInput.trim().replace(/^=/, "").trim();
  if (!normalized.startsWith("{") || !normalized.endsWith("}")) {
    return [];
  }
  const body = normalized.slice(1, -1);
  const values: string[] = [];
  let current = "";
  let inString = false;
  for (let index = 0; index < body.length; index += 1) {
    const ch = body[index];
    if (ch === '"') {
      if (inString && body[index + 1] === '"') {
        current += '"';
        index += 1;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (!inString && (ch === ";" || ch === ",")) {
      const token = current.trim();
      if (token) {
        values.push(token);
      }
      current = "";
      continue;
    }
    current += ch;
  }
  const finalToken = current.trim();
  if (finalToken) {
    values.push(finalToken);
  }
  return values;
};

const parseFunctionArgs = (value: string): string[] =>
  value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length > 0);

export const buildLambdaFormula = (bodyInput: string, argsInput: string): string => {
  const normalizedBody = bodyInput.trim();
  if (!normalizedBody) {
    throw new Error("Function definition is required.");
  }

  const bodyWithoutEquals = normalizedBody.replace(/^=/, "").trim();
  if (/^LAMBDA\s*\(/i.test(bodyWithoutEquals)) {
    return `=${bodyWithoutEquals}`;
  }

  const args = parseFunctionArgs(argsInput);
  const seen = new Set<string>();
  args.forEach((arg) => {
    if (!IDENTIFIER_PATTERN.test(arg)) {
      throw new Error(`Invalid function argument "${arg}".`);
    }
    const key = arg.toUpperCase();
    if (seen.has(key)) {
      throw new Error(`Duplicate function argument "${arg}".`);
    }
    seen.add(key);
  });

  return args.length > 0
    ? `=LAMBDA(${args.join(",")},${bodyWithoutEquals})`
    : `=LAMBDA(${bodyWithoutEquals})`;
};

export const buildNamedListFormula = (listInput: string): string => {
  const values = listInput
    .split(/[\r\n,;]+/)
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
  if (values.length === 0) {
    throw new Error("List values are required.");
  }

  const literals = values.map((value) => {
    if (/^-?\d+(\.\d+)?$/.test(value)) {
      return value;
    }
    if (/^(TRUE|FALSE)$/i.test(value)) {
      return value.toUpperCase();
    }
    return `"${value.replace(/"/g, '""')}"`;
  });

  return `={${literals.join(";")}}`;
};

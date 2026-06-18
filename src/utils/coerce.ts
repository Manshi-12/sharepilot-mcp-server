import { ColumnInfo } from "./resolve.js";

/**
 * Coerces a raw user-supplied value into the shape SharePoint expects for a
 * given column type. Throws a descriptive error for invalid choice values so
 * the caller can surface it as a fieldError rather than letting SharePoint
 * silently accept bad data.
 */
export function coerceValue(col: ColumnInfo, value: any): any {
  switch (col.type) {
    case "boolean":
      if (typeof value === "boolean") return value;
      return String(value).trim().toLowerCase() === "true";

    case "number":
    case "currency": {
      const n = typeof value === "number" ? value : Number(value);
      if (isNaN(n)) {
        throw new Error(
          `"${value}" is not a valid number for column "${col.displayName}".`
        );
      }
      return n;
    }

    case "dateTime": {
      const d = new Date(value);
      if (isNaN(d.getTime())) {
        throw new Error(
          `"${value}" is not a valid date/time for column "${col.displayName}". ` +
          `Use a format like "2025-06-17" or "2025-06-17T10:00:00Z".`
        );
      }
      return d.toISOString();
    }

    case "choice": {
      const strVal = String(value).trim();
      if (col.choices && col.choices.length > 0) {
        const match = col.choices.find(
          (c) => c.toLowerCase() === strVal.toLowerCase()
        );
        if (!match) {
          throw new Error(
            `"${strVal}" is not a valid choice for column "${col.displayName}". ` +
            `Valid options are: ${col.choices.join(", ")}.`
          );
        }
        return match;
      }
      return strVal;
    }

    case "multiChoice": {
      const arr = Array.isArray(value) ? value : [value];
      if (col.choices && col.choices.length > 0) {
        const resolved: string[] = [];
        for (const v of arr) {
          const strVal = String(v).trim();
          const match = col.choices.find(
            (c) => c.toLowerCase() === strVal.toLowerCase()
          );
          if (!match) {
            throw new Error(
              `"${strVal}" is not a valid choice for column "${col.displayName}". ` +
              `Valid options are: ${col.choices.join(", ")}.`
            );
          }
          resolved.push(match);
        }
        return resolved;
      }
      return arr;
    }

    case "hyperlinkOrPicture": {
      // Graph requires the exact capitalized keys "Url" and "Description" —
      // accept a plain URL string or an object in any casing the caller used
      // (url/Url/URL, description/Description/label) and normalize it.
      if (typeof value === "string") {
        return { Url: value, Description: value };
      }
      if (value && typeof value === "object") {
        const url = value.Url || value.url || value.URL;
        const description = value.Description || value.description || value.label || value.Label || url;
        if (!url) {
          throw new Error(
            `Column "${col.displayName}" needs a URL — pass a link string, or an object with a "url" property.`
          );
        }
        return { Url: url, Description: description };
      }
      throw new Error(
        `Column "${col.displayName}" needs a URL — pass a link string, or an object with a "url" property.`
      );
    }

    default:
      return value;
  }
}
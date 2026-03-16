/**
 * Type-safe DOM element getter. Throws a descriptive error if the element
 * is not found instead of silently failing with null/undefined at runtime.
 */
export function getElement<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Required DOM element #${id} not found`);
  return el as T;
}

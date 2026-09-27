import { attrOf, child, children, OoxmlNode } from './ooxml.js';

export interface Relationship {
  target: string;
  external: boolean;
  /** The relationship Type URI ('' when the file omits it). */
  type: string;
}

/** Parse a `_rels/*.rels` part into a map of relationship id → target. */
export function buildRels(
  root: OoxmlNode | undefined,
): Map<string, Relationship> {
  const map = new Map<string, Relationship>();
  for (const rel of children(child(root, 'Relationships'), 'Relationship')) {
    const id = attrOf(rel, 'Id');
    const target = attrOf(rel, 'Target');
    if (id && target) {
      map.set(id, {
        target,
        external: attrOf(rel, 'TargetMode') === 'External',
        type: attrOf(rel, 'Type') ?? '',
      });
    }
  }
  return map;
}

import { CellAddress } from '../model/types';

function addrKey(addr: CellAddress): string {
  return `${addr.sheet}!${addr.col},${addr.row}`;
}

export class DependencyGraph {
  /** cell -> set of cells it depends on (precedents) */
  private precedents: Map<string, Set<string>> = new Map();
  /** cell -> set of cells that depend on it (dependents) */
  private dependents: Map<string, Set<string>> = new Map();

  setDependencies(cell: CellAddress, dependsOn: CellAddress[]): void {
    const cellKey = addrKey(cell);

    // Remove old dependencies
    this.removeDependencies(cell);

    // Set new dependencies
    const precSet = new Set<string>();
    for (const dep of dependsOn) {
      const depKey = addrKey(dep);
      precSet.add(depKey);

      // Update reverse map
      let depSet = this.dependents.get(depKey);
      if (!depSet) {
        depSet = new Set();
        this.dependents.set(depKey, depSet);
      }
      depSet.add(cellKey);
    }

    if (precSet.size > 0) {
      this.precedents.set(cellKey, precSet);
    }
  }

  removeDependencies(cell: CellAddress): void {
    const cellKey = addrKey(cell);
    const oldPrec = this.precedents.get(cellKey);
    if (oldPrec) {
      for (const depKey of oldPrec) {
        const depSet = this.dependents.get(depKey);
        if (depSet) {
          depSet.delete(cellKey);
          if (depSet.size === 0) this.dependents.delete(depKey);
        }
      }
      this.precedents.delete(cellKey);
    }
  }

  /**
   * Given a set of changed cells, return all cells that need recalculation
   * in topological order (dependencies evaluated before dependents),
   * plus the set of cells involved in circular references.
   */
  getRecalcOrder(changedCells: CellAddress[]): { order: CellAddress[]; cyclic: CellAddress[] } {
    // BFS to find all affected cells
    const affected = new Set<string>();
    const queue: string[] = [];

    for (const cell of changedCells) {
      const key = addrKey(cell);
      queue.push(key);
      affected.add(key);
    }

    while (queue.length > 0) {
      const key = queue.shift()!;
      const deps = this.dependents.get(key);
      if (deps) {
        for (const depKey of deps) {
          if (!affected.has(depKey)) {
            affected.add(depKey);
            queue.push(depKey);
          }
        }
      }
    }

    // Topological sort using Kahn's algorithm
    // Build in-degree map restricted to affected cells
    const inDegree = new Map<string, number>();
    const adjList = new Map<string, Set<string>>();

    for (const key of affected) {
      inDegree.set(key, 0);
      adjList.set(key, new Set());
    }

    for (const key of affected) {
      const prec = this.precedents.get(key);
      if (prec) {
        for (const precKey of prec) {
          if (affected.has(precKey)) {
            inDegree.set(key, (inDegree.get(key) ?? 0) + 1);
            adjList.get(precKey)!.add(key);
          }
        }
      }
    }

    // Start with nodes that have 0 in-degree
    const sorted: string[] = [];
    const zeroQueue: string[] = [];

    for (const [key, deg] of inDegree) {
      if (deg === 0) zeroQueue.push(key);
    }

    while (zeroQueue.length > 0) {
      const key = zeroQueue.shift()!;
      sorted.push(key);
      const neighbors = adjList.get(key);
      if (neighbors) {
        for (const nKey of neighbors) {
          const newDeg = (inDegree.get(nKey) ?? 1) - 1;
          inDegree.set(nKey, newDeg);
          if (newDeg === 0) zeroQueue.push(nKey);
        }
      }
    }

    // Nodes not in sorted are part of a cycle — collect them separately
    const sortedSet = new Set(sorted);
    const cyclicKeys: string[] = [];
    for (const key of affected) {
      if (!sortedSet.has(key)) {
        cyclicKeys.push(key);
      }
    }

    return {
      order: sorted.map(this.parseKey),
      cyclic: cyclicKeys.map(this.parseKey),
    };
  }

  detectCycle(startCell: CellAddress): CellAddress[] | null {
    const startKey = addrKey(startCell);
    const visited = new Set<string>();
    const path: string[] = [];
    const pathSet = new Set<string>();

    const dfs = (key: string): boolean => {
      if (pathSet.has(key)) {
        // Found cycle
        return true;
      }
      if (visited.has(key)) return false;
      visited.add(key);
      path.push(key);
      pathSet.add(key);

      const prec = this.precedents.get(key);
      if (prec) {
        for (const precKey of prec) {
          if (dfs(precKey)) return true;
        }
      }

      path.pop();
      pathSet.delete(key);
      return false;
    };

    if (dfs(startKey)) {
      return path.map(this.parseKey);
    }
    return null;
  }

  private parseKey(key: string): CellAddress {
    const excl = key.indexOf('!');
    const sheet = key.substring(0, excl);
    const rest = key.substring(excl + 1);
    const comma = rest.indexOf(',');
    return {
      sheet,
      col: parseInt(rest.substring(0, comma), 10),
      row: parseInt(rest.substring(comma + 1), 10),
    };
  }

  clear(): void {
    this.precedents.clear();
    this.dependents.clear();
  }
}

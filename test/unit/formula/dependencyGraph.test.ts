import { describe, it, expect } from 'vitest';
import { DependencyGraph } from '../../../src/formula/DependencyGraph';

describe('DependencyGraph', () => {
  it('tracks dependencies', () => {
    const graph = new DependencyGraph();
    graph.setDependencies(
      { sheet: 'S1', col: 2, row: 0 }, // C1
      [
        { sheet: 'S1', col: 0, row: 0 },
        { sheet: 'S1', col: 1, row: 0 },
      ], // depends on A1, B1
    );

    const order = graph.getRecalcOrder([{ sheet: 'S1', col: 0, row: 0 }]); // A1 changed
    const keys = order.map((a) => `${a.col},${a.row}`);
    expect(keys).toContain('0,0'); // A1
    expect(keys).toContain('2,0'); // C1 (dependent)
    // C1 should come after A1
    expect(keys.indexOf('2,0')).toBeGreaterThan(keys.indexOf('0,0'));
  });

  it('handles chain dependencies', () => {
    const graph = new DependencyGraph();
    // B1 depends on A1, C1 depends on B1
    graph.setDependencies({ sheet: 'S1', col: 1, row: 0 }, [{ sheet: 'S1', col: 0, row: 0 }]);
    graph.setDependencies({ sheet: 'S1', col: 2, row: 0 }, [{ sheet: 'S1', col: 1, row: 0 }]);

    const order = graph.getRecalcOrder([{ sheet: 'S1', col: 0, row: 0 }]);
    const keys = order.map((a) => `${a.col},${a.row}`);
    expect(keys.indexOf('1,0')).toBeGreaterThan(keys.indexOf('0,0'));
    expect(keys.indexOf('2,0')).toBeGreaterThan(keys.indexOf('1,0'));
  });

  it('removes dependencies', () => {
    const graph = new DependencyGraph();
    graph.setDependencies({ sheet: 'S1', col: 1, row: 0 }, [{ sheet: 'S1', col: 0, row: 0 }]);
    graph.removeDependencies({ sheet: 'S1', col: 1, row: 0 });

    const order = graph.getRecalcOrder([{ sheet: 'S1', col: 0, row: 0 }]);
    // Only A1 itself, B1 no longer depends on it
    expect(order.length).toBe(1);
  });

  it('detects cycles', () => {
    const graph = new DependencyGraph();
    graph.setDependencies({ sheet: 'S1', col: 0, row: 0 }, [{ sheet: 'S1', col: 1, row: 0 }]);
    graph.setDependencies({ sheet: 'S1', col: 1, row: 0 }, [{ sheet: 'S1', col: 0, row: 0 }]);

    const cycle = graph.detectCycle({ sheet: 'S1', col: 0, row: 0 });
    expect(cycle).not.toBeNull();
  });
});

import { NodeViewWrapper, type NodeViewProps } from '@tiptap/react';
import { useEffect, useRef, useState } from 'react';
import { getLinkedRange, subscribeLinkedRanges } from '../../data/linkedRanges';
import { useReducedMotion } from '../../chrome/ui';
import { tableFromRange, tableSignature, type LinkedTableModel } from './linkedTable';

export function LinkedTableView({ node, deleteNode }: NodeViewProps) {
  const linkId = String(node.attrs.linkId ?? '');
  const reduced = useReducedMotion();
  const [model, setModel] = useState<LinkedTableModel | null>(null);
  const [missing, setMissing] = useState(false);
  const [flowing, setFlowing] = useState(false);
  const previous = useRef('');

  useEffect(() => {
    let stop = false;
    async function load(animate: boolean) {
      const range = await getLinkedRange(linkId);
      if (stop) return;
      if (!range) {
        setMissing(true);
        setModel(null);
        return;
      }
      const next = tableFromRange(range);
      const signature = tableSignature(next.rows);
      if (animate && previous.current && previous.current !== signature && !reduced) {
        setFlowing(true);
        window.setTimeout(() => setFlowing(false), 700);
      }
      previous.current = signature;
      setMissing(false);
      setModel(next);
    }
    void load(false);
    const unsubscribe = subscribeLinkedRanges((range) => {
      if (range.id === linkId) void load(true);
    });
    const timer = window.setInterval(() => void load(true), 2000);
    return () => {
      stop = true;
      unsubscribe();
      window.clearInterval(timer);
    };
  }, [linkId, reduced]);

  return (
    <NodeViewWrapper className={`linked-table${flowing ? ' is-flowing' : ''}`} data-linked-table={linkId}>
      <header>
        <span>{model ? model.caption : 'Linked sheet range'}</span>
        <button type="button" onClick={deleteNode}>Remove</button>
      </header>
      {missing ? <p>This range is not on the desk.</p> : null}
      {model && model.rows.length === 0 ? <p>The range is empty.</p> : null}
      {model && model.rows.length > 0 ? (
        <table>
          <tbody>
            {model.rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, colIndex) => rowIndex === 0
                  ? <th key={colIndex}>{cell}</th>
                  : <td key={colIndex}>{cell}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      ) : null}
    </NodeViewWrapper>
  );
}

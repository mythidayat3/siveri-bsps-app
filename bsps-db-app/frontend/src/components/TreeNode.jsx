import { useState } from 'react';

export default function TreeNode({ node, level = 0 }) {
  const [isExpanded, setIsExpanded] = useState(false);
  const hasChildren = node.children && node.children.length > 0;

  return (
    <div style={{ display: 'flex', flexDirection: 'column' }}>
      <div 
        className={`tree-row ${isExpanded ? 'active-row' : ''}`}
        style={{ paddingLeft: `${14 + level * 16}px` }}
      >
        <div 
          className="tree-node-title" 
          onClick={() => hasChildren && setIsExpanded(!isExpanded)}
        >
          {hasChildren ? (
            <span className={`tree-arrow ${isExpanded ? 'expanded' : ''}`} style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '20px', height: '20px' }}>
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="9 18 15 12 9 6"></polyline>
              </svg>
            </span>
          ) : (
            <span style={{ width: '20px', display: 'inline-block' }}></span>
          )}
          <span>{node.name}</span>
        </div>
        <div className="tree-stats">
          <span className="tree-badge" title="Total Calon Penerima Bantuan">CPB: {node.cpb}</span>
          <span className="tree-badge green" title="Lolos Verifikasi">Lolos: {node.lolos}</span>
          <span className="tree-badge yellow" title="Tidak Lolos Verifikasi">T.Lolos: {node.tidak_lolos}</span>
          <span className="tree-badge red" title="Belum Terverifikasi">Belum: {node.belum_verifikasi}</span>
        </div>
      </div>
      {hasChildren && isExpanded && (
        <div className="tree-children" style={{ marginLeft: `${level === 0 ? 12 : 20}px` }}>
          {node.children.map((child, idx) => (
            <TreeNode key={idx} node={child} level={level + 1} />
          ))}
        </div>
      )}
    </div>
  );
}

import fs from 'node:fs/promises';
const icons={
  toolbox:['#087F8C','<rect x="4" y="7" width="16" height="14" rx="3"/><path d="M8 7V4h8v3M4 12h16M10 11v4h4v-4"/>'],
  heading1:['#3660BD','<path d="M4 5v14M13 5v14M4 12h9M17 9l3-2v12M17 19h6"/>'],
  heading2:['#3660BD','<path d="M3 5v14M11 5v14M3 12h8M16 10c0-4 6-4 6 0 0 3-6 5-6 9h6"/>'],
  paragraph:['#3660BD','<path d="M4 5h16M4 10h16M4 15h12M4 20h16"/>'],
  list:['#3660BD','<path d="M9 5h12M9 12h12M9 19h12M3 5h1M3 12h1M3 19h1"/>'],
  table:['#16846C','<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M3 15h18M9 4v16M15 4v16"/>'],
  clean:['#A97720','<path d="M5 5h10M5 10h10M5 15h6M5 20h6M16 13l5 5M21 13l-5 5"/>'],
  replace:['#087F8C','<path d="M4 7h14l-3-3M20 17H6l3 3M6 4L3 10M18 14l3 6"/>'],
  header:['#16846C','<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M10 10v10"/><rect x="5" y="6" width="14" height="2" rx="1" fill="#16846C" stroke="none"/>'],
  columns:['#16846C','<path d="M7 4v16M17 4v16M3 12h18M5 9l-3 3 3 3M19 9l3 3-3 3"/>'],
  wrap:['#16846C','<path d="M3 5h18M3 11h14a4 4 0 010 8h-5M3 17h4M14 16l-3 3 3 3"/>'],
  filter:['#16846C','<path d="M3 4h18l-7 8v7l-4 2v-9z"/>'],
  trim:['#A97720','<path d="M4 4v16M20 4v16M9 8h6M9 12h6M9 16h6M1 12h5M18 12h5"/>'],
  presentation:['#8560B5','<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M12 17v4M8 21h8M7 8h10M7 12h6"/>'],
  type:['#8560B5','<path d="M4 4h16M12 4v16M8 20h8M4 4v3M20 4v3"/>'],
  duplicate:['#8560B5','<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M16 4H5a2 2 0 00-2 2v10M11 14h7M14 11v7"/>'],
  capture:['#087F8C','<path d="M8 3H4v4M16 3h4v4M4 17v4h4M20 17v4h-4"/><rect x="8" y="8" width="8" height="8" rx="1"/>'],
  insert:['#087F8C','<path d="M5 4h8M5 4v16h14v-6M13 10h9M18 6l4 4-4 4"/>'],
  pdf:['#C05C55','<path d="M14 3H6a2 2 0 00-2 2v14a2 2 0 002 2h12a2 2 0 002-2V9zM14 3v6h6M8 13h8M8 17h5"/>'],
  arrow:['#087F8C','<path d="M4 12h16M14 6l6 6-6 6"/>'],
  document:['#3660BD','<path d="M14 3H6a2 2 0 00-2 2v14a2 2 0 002 2h12a2 2 0 002-2V9zM14 3v6h6M8 13h8M8 17h8"/>'],
  check:['#16846C','<circle cx="12" cy="12" r="9"/><path d="M7 12l3 3 7-7"/>'],
  undo:['#526D7C','<path d="M7 5L3 9l4 4M3 9h11a6 6 0 010 12"/>'],
  copy:['#526D7C','<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M16 4H5a2 2 0 00-2 2v10"/>'],
  up:['#526D7C','<path d="M12 20V4M5 11l7-7 7 7"/>'],
  down:['#526D7C','<path d="M12 4v16M5 13l7 7 7-7"/>'],
  plus:['#087F8C','<path d="M12 4v16M4 12h16"/>'],
  match:['#087F8C','<rect x="4" y="3" width="14" height="6" rx="1.5"/><path d="M18 6h2v5h-8v3M11 14h2v7h-2z"/>'],
  cards:['#087F8C','<rect x="7" y="3" width="13" height="10" rx="2"/><rect x="4" y="9" width="13" height="10" rx="2" fill="#fff"/><path d="M7 13h7M7 16h4"/>'],
  dedupe:['#16846C','<rect x="3" y="4" width="11" height="4" rx="1"/><rect x="3" y="11" width="11" height="4" rx="1"/><path d="M17 9l4 4M21 9l-4 4M3 19h11"/>'],
  backup:['#526D7C','<path d="M4 7a2 2 0 012-2h9l5 5v8a2 2 0 01-2 2H6a2 2 0 01-2-2z"/><path d="M12 17v-6M9.5 13.5L12 11l2.5 2.5"/>'],
  trash:['#A76555','<path d="M3 6h18M9 6V3h6v3M6 6l1 15h10l1-15M10 10v7M14 10v7"/>']
};
const dir=new URL('../public/icons/',import.meta.url);await fs.mkdir(dir,{recursive:true});
for(const [name,[color,shapes]] of Object.entries(icons))await fs.writeFile(new URL(name+'.svg',dir),`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">${shapes}</svg>`);
console.log(`Generated ${Object.keys(icons).length} original SVG icons.`);

import type { CepLevel } from './CepApiModels';

/**
 * Level presentation mapping for UI text and icons.
 */
export function getLevelLabel(level: CepLevel | string): string {
  if (level === 'Explorer') return 'Explorer';
  if (level === 'Practitioner') return 'Practitioner';
  if (level === 'Master') return 'Master';
  return String(level ?? '');
}

export function getLevelIcon(level: CepLevel | string): string {
  if (level === 'Explorer') return '🧭';
  if (level === 'Practitioner') return '🛠️';
  if (level === 'Master') return '🏆';
  return '⭐';
}

export function getNextLevelLabel(level: CepLevel | string): string {
  if (level === 'Explorer') return 'Practitioner';
  if (level === 'Practitioner') return 'Master';
  return '';
}

export function isTopLevel(level: CepLevel | string): boolean {
  return level === 'Master';
}

import type { CepLevel } from './CepApiModels';

/**
 * Level presentation mapping for UI text and icons.
 * Supports both legacy and new level values.
 */
export function getLevelLabel(level: CepLevel | string): string {
  if (level === 'Bronze' || level === 'Explorer') return 'Explorer';
  if (level === 'Silver' || level === 'Practitioner') return 'Practitioner';
  if (level === 'Gold' || level === 'Master') return 'Master';
  return String(level ?? '');
}

export function getLevelIcon(level: CepLevel | string): string {
  if (level === 'Bronze' || level === 'Explorer') return '🧭';
  if (level === 'Silver' || level === 'Practitioner') return '🛠️';
  if (level === 'Gold' || level === 'Master') return '🏆';
  return '⭐';
}

export function getNextLevelLabel(level: CepLevel | string): string {
  if (level === 'Bronze' || level === 'Explorer') return 'Practitioner';
  if (level === 'Silver' || level === 'Practitioner') return 'Master';
  return '';
}

export function isTopLevel(level: CepLevel | string): boolean {
  return level === 'Gold' || level === 'Master';
}

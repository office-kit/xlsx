// Formulas ▸ Recently Used. Excel seeds a fresh profile with this list and
// moves each function the user inserts to the front, keeping ten.

const LIMIT = 10;

class RecentFunctions {
  names = $state(['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV']);

  use(name: string): void {
    this.names = [name, ...this.names.filter((n) => n !== name)].slice(0, LIMIT);
  }
}

export const recentFunctions = new RecentFunctions();

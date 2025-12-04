/**
 * Basic test to ensure Jest is working
 */

describe('fixEcalendar', () => {
  it('should pass basic test', () => {
    expect(true).toBe(true);
  });

  it('should have correct package structure', () => {
    const pkg = require('../../package.json');
    expect(pkg.name).toBe('fixecalendar');
    expect(pkg.version).toBeTruthy();
  });
});

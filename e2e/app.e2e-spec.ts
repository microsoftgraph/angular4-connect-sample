import { DemoSimplePage } from './app.po';

describe('demo-simple App', () => {
  let page: DemoSimplePage;

  beforeEach(() => {
    page = new DemoSimplePage();
  });

  it('should display welcome message', () => {
    page.navigateTo();
    expect(page.getParagraphText()).toEqual('Welcome to app!');
  });
});

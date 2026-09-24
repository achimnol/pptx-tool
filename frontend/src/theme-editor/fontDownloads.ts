export interface FontDownload {
  name: string;
  href: string;
}

/** The official download pages of the fonts used by the bundled themes. */
export const FONT_DOWNLOADS: readonly FontDownload[] = [
  {name: 'Pretendard', href: 'https://cactus.tistory.com/306'},
  {name: 'SNU Appendard', href: 'https://qbio.io/share/fonts/'},
  {name: 'Inter and Inter Display', href: 'https://fonts.google.com/specimen/Inter'},
  {name: 'Sarasa Term K', href: 'https://picaq.github.io/sarasa'},
  {name: 'Paperlogy', href: 'https://freesentation.blog/paperlogyfont'},
  {name: 'Freesentation', href: 'https://freesentation.blog/freesentation'},
  {name: 'Nanum Fonts', href: 'https://hangeul.naver.com/font'},
];

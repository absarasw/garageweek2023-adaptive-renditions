/* global WebImporter */
/* eslint-disable no-console, class-methods-use-this */

const reorganiseHero = (main, document) => {
  const heading = document.querySelector('.article-title__StyledArticleTitle-sc-1rce01p-0');
  if (heading) {
    const h1 = document.createElement('h1');
    h1.innerHTML = heading.innerHTML;
    heading.replaceWith(h1);
  }
};

const createMetadata = (main, url, document) => {
  const meta = {};

  const title = document.querySelector('title');
  if (title) {
    meta.Title = title.innerHTML.replace(/[\n\t]/gm, '');
  }

  const desc = document.querySelector('[property="og:description"]');
  if (desc) {
    meta.Description = desc.content;
  }

  const img = document.querySelector('[property="og:image"]');
  if (img) {
    const el = document.createElement('img');
    el.src = img.content;

    const alt = document.querySelector('[property="og:image:alt"], [property="twitter:image:alt"]');
    if (alt) {
      el.alt = alt.content;
    }
    meta.Image = el;
  }

  const subtitle = document.querySelector('.article-header__Topline-sc-1f6j4tn-2');
  if (subtitle) {
    meta.Subtitle = subtitle.textContent;
  }

  const date = document.querySelector('.article-header__Date-sc-1f6j4tn-3');
  if (date) {
    meta.Date = date.textContent.replace('Uhr', '').trim();
  } else {
    const m = url.match(/(\d{1,2})\/(\d{1,2})\//);
    if (m && m.length > 2) {
      meta.Date = `01.${m[2]}.${m[1]}, 09:00`;
    }
  }

  const tags = document.querySelector('.meta-section__Tags-sc-m65hdj-4');
  if (tags) {
    const array = [];
    tags.querySelectorAll('span').forEach((span) => {
      array.push(span.textContent);
    });
    meta.Tags = array.join(', ');
  }

  const block = WebImporter.Blocks.getMetadataBlock(document, meta);
  main.append(block);

  return meta;
};

const makeAbsoluteLinks = (main, host, base) => {
  main.querySelectorAll('a').forEach((a) => {
    if (a.href.startsWith('/')) {
      const ori = a.href;
      const u = new URL(a.href, host);
      if (base && u.pathname.startsWith(base)) {
        u.pathname = u.pathname.substring(base.length);
      }
      u.pathname = u.pathname.replace(/\.html$/, '').toLocaleLowerCase();
      a.href = u.toString();

      if (a.textContent === ori) {
        a.textContent = a.href;
      }
    }
  });
};

const createVideoBlock = (main, host, document) => {
  main.querySelectorAll('a.video-description__VideoLink-sc-18c4yrr-3').forEach((a) => {
    if (a.href) {
      const parent = a.closest('.video-paywall__VideoLayout-sc-55w0el-1');
      const u = new URL(a.href, host);
      a.href = u.toString();
      a.innerHTML = u.toString();
      const cells = [['Video'], [a]];
      const table = WebImporter.DOMUtils.createTable(cells, document);
      parent.replaceWith(table);
    }
  });
};

const createTeaserBlock = (main, document) => {
  main.querySelectorAll('a.teaser__TeaserLink-sc-44nuea-0').forEach((a) => {
    if (a.href) {
      const parent = a.parentNode;
      a.innerHTML = a.href;
      const cells = [['Teaser'], [a]];
      const table = WebImporter.DOMUtils.createTable(cells, document);
      parent.append(table);
    }
  });
};

const createCaption = (main, document) => {
  main.querySelectorAll('figcaption').forEach((caption) => {
    const p = document.createElement('p');
    const i = document.createElement('i');
    i.innerHTML = caption.innerHTML;
    p.append(i);
    caption.replaceWith(p);
  });
};

const createQuoteBlock = (main, document) => {
  main.querySelectorAll('blockquote').forEach((quote) => {
    const div = document.createElement('div');
    const p = quote.querySelector('p');
    if (p) {
      div.append(p);
    }
    const footer = quote.querySelector('footer');
    if (footer) {
      const h = document.createElement('h3');
      h.innerHTML = footer.innerHTML;
      div.append(h);
    }
    const cells = [['Quote'], [div]];
    const table = WebImporter.DOMUtils.createTable(cells, document);
    quote.replaceWith(table);
  });
};

const createEmblemsBlock = (main, document) => {
  main.querySelectorAll('.emblems__Layout-sc-176j480-1').forEach((emblem) => {
    const parent = emblem.parentNode;
    const div = document.createElement('div');
    const cells = [['Emblems'], [div]];
    emblem.querySelectorAll('img').forEach((img) => {
      const p = document.createElement('p');
      p.innerHTML = `:${img.alt}:`;
      div.append(p);
    });
    const table = WebImporter.DOMUtils.createTable(cells, document);
    parent.replaceWith(table);
  });
};

export default {
  /**
   * Apply DOM operations to the provided document and return
   * the root element to be then transformed to Markdown.
   * @param {HTMLDocument} document The document
   * @returns {HTMLElement} The root element
   */
  transformDOM: async ({ document, url }) => {
    const host = 'https://fcbayern.com';
    const main = document.querySelector('.layout__PageWrapperLayout-sc-qwcvwt-0');

    reorganiseHero(main, document);
    // createRelatedStoriesBlock(main, document);
    createMetadata(main, url, document);
    createVideoBlock(main, host, document);
    createTeaserBlock(main, document);
    createCaption(main, document);
    createQuoteBlock(main, document);
    createEmblemsBlock(main, document);

    WebImporter.DOMUtils.remove(main, [
      'aside',
      'header',
      '.article-header__ArticleMetaData-sc-1f6j4tn-1',
      '.article-page__ActionWrapper-sc-1zjme2-3',
      '.meta-section__Layout-sc-m65hdj-0',
      '.bskjoz',
      '.webradio-sticky-player__WebradioPlayerRoot-sc-3y91hb-6',
      '.page__Content-sc-afw92i-1',
      '.page-header__PageHeaderAppView-sc-1akn7gm-0',
      '.match-navigation-bar__Layout-sc-3z74jw-0',
      '.match-stage-hero__Layout-sc-1i5avbw-0',
      '.match-info__Layout-sc-rptmyj-0',
    ]);

    makeAbsoluteLinks(main, 'https://main--fcbayern--hlxsites.hlx.page');

    return main;
  },

  /**
   * Return a path that describes the document being transformed (file name, nesting...).
   * The path is then used to create the corresponding Word document.
   * @param {String} url The url of the document being transformed.
   * @param {HTMLDocument} document The document
   */
  // eslint-disable-next-line arrow-body-style
  generateDocumentPath: ({ url }) => {
    return new URL(url).pathname.replace(/\.html$/, '').toLocaleLowerCase();
  },
};

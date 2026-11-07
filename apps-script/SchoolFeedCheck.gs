const SCHOOL_FEED_CHECK_URLS = {
  news: 'https://mms.usd266.com/fs/post-manager/boards/52/posts/feed.rss',
  social: 'https://www.juicer.io/api/feeds/378904?filter=783230,778997',
  calendar: 'https://mms.usd266.com/calendar/calendar.ics?calendar_ids=8,7,34,17'
};

function testSchoolFeeds() {
  const failures = [];
  Object.keys(SCHOOL_FEED_CHECK_URLS).forEach(name => {
    try {
      const response = UrlFetchApp.fetch(SCHOOL_FEED_CHECK_URLS[name], {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          Accept: 'application/rss+xml, application/xml, application/json, text/calendar, text/plain, */*'
        }
      });
      const status = response.getResponseCode();
      if (status < 200 || status >= 300) {
        throw new Error('HTTP ' + status + '. The server did not allow a successful feed request.');
      }
      const body = response.getContentText();
      const sample = schoolFeedCheckSample_(name, body);
      console.log(JSON.stringify({ feed: name, status: status, sample: sample }, null, 2));
    } catch (err) {
      failures.push(name + ': ' + err.message);
      console.error('School feed check failed — ' + name + ': ' + err.message);
    }
  });
  if (failures.length) {
    try {
      const response = UrlFetchApp.fetch('https://mms.usd266.com/', {
        muteHttpExceptions: true,
        followRedirects: true
      });
      const status = response.getResponseCode();
      if (status !== 200) throw new Error('School homepage returned HTTP ' + status);
      const html = response.getContentText();
      const candidates = [...html.matchAll(/\b(?:href|src)\s*=\s*["']([^"']+)["']/gi)]
        .map(match => match[1].replace(/&amp;/g, '&'))
        .filter(url => /rss|calendar|feed|news/i.test(url));
      console.log(JSON.stringify({
        homepage_status: status,
        possible_feed_or_page_links: [...new Set(candidates)].slice(0, 30)
      }, null, 2));
    } catch (err) {
      console.error('School homepage check failed: ' + err.message);
    }
  }
  if (failures.length) {
    throw new Error('Some feeds could not be read. Successful samples are in the execution log. ' + failures.join('; '));
  }
  console.log('All three feeds returned recognizable data. No spreadsheet or website content was changed.');
}

function schoolFeedCheckSample_(name, body) {
  if (name === 'news') {
    const root = XmlService.parse(body).getRootElement();
    let entries;
    let namespace;
    if (root.getName() === 'rss') {
      const channel = root.getChild('channel');
      if (!channel) throw new Error('RSS response has no channel.');
      entries = channel.getChildren('item');
    } else if (root.getName() === 'feed') {
      namespace = root.getNamespace();
      entries = root.getChildren('entry', namespace);
    } else {
      throw new Error('Expected RSS or Atom XML, not ' + root.getName() + '.');
    }
    return {
      count: entries.length,
      items: entries.slice(0, 3).map(entry => {
        const text = key => {
          const child = namespace ? entry.getChild(key, namespace) : entry.getChild(key);
          return child ? child.getText() : '';
        };
        return {
          title: schoolFeedCheckText_(text('title')),
          description: schoolFeedCheckText_(text('description') || text('summary') || text('content')).slice(0, 1500)
        };
      })
    };
  }
  if (name === 'calendar') {
    const unfolded = body.replace(/\r?\n[ \t]/g, '');
    if (!/^BEGIN:VCALENDAR\s*$/m.test(unfolded) || !/^END:VCALENDAR\s*$/m.test(unfolded)) {
      throw new Error('Expected an iCalendar response.');
    }
    const events = [...unfolded.matchAll(/^BEGIN:VEVENT\r?\n([\s\S]*?)^END:VEVENT\s*$/gm)];
    return {
      count: events.length,
      items: events.slice(0, 3).map(event => {
        const text = key => {
          const match = event[1].match(new RegExp('^' + key + '(?:;[^:\\r\\n]*)?:(.*)$', 'm'));
          return match ? match[1].trim().replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1') : '';
        };
        return { title: text('SUMMARY'), starts: text('DTSTART'), description: text('DESCRIPTION').slice(0, 1500) };
      })
    };
  }
  const data = JSON.parse(body);
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error('Expected a Juicer JSON object.');
  }
  if (!data.posts || !Array.isArray(data.posts.items)) {
    throw new Error('Juicer returned JSON without the expected posts.items array. Top-level keys: ' + Object.keys(data).join(', '));
  }
  return {
    count: data.posts.items.length,
    post_keys: data.posts.items.length ? Object.keys(data.posts.items[0]) : [],
    posts: data.posts.items.slice(0, 3).map(post => ({
      id: post.id,
      source: post.source,
      external: post.external,
      text: schoolFeedCheckText_(post.message || post.text || post.description || '').slice(0, 1500),
      created_at: post.created_at
    }))
  };
}

function schoolFeedCheckText_(value) {
  return String(value)
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function inspectSocialPost() {
  const url = 'https://www.juicer.io/api/feeds/378904?filter=783230,778997';
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const status = response.getResponseCode();
    if (status !== 200) throw new Error('HTTP ' + status);
    const data = JSON.parse(response.getContentText());
    if (!data.posts || !Array.isArray(data.posts.items)) {
      throw new Error('Response has no posts.items array.');
    }
    const posts = data.posts.items;
    console.log(JSON.stringify({
      count: posts.length,
      first_post: posts.length ? posts[0] : null
    }, null, 2).slice(0, 10000));
  } catch (err) {
    console.error('Social check failed: ' + err.message);
  }
}

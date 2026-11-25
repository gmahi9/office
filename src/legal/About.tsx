import { thirdPartyNotices } from './notices';

export const SOURCE_URL = 'https://github.com/gmahi9/office';

export function About({ onClose }: { onClose: () => void }) {
  return (
    <section className="about" aria-labelledby="about-title">
      <header className="dialog-head">
        <h2 id="about-title">Licences</h2>
        <button type="button" className="text-btn" onClick={onClose}>Close</button>
      </header>
      <p>
        Office is free software. You can redistribute it and/or modify it under the terms of the
        GNU General Public License as published by the Free Software Foundation, either version 3
        of the License, or (at your option) any later version.
      </p>
      <p>
        The suite as distributed includes HyperFormula, which is GPL-3.0-only. Because of that,
        this combined program is distributed under GPL-3.0. There is no warranty. See the
        GNU General Public License for details.
      </p>
      <p>
        Source: <a href={SOURCE_URL}>{SOURCE_URL}</a>
      </p>
      <h3>Third-party notices</h3>
      <ul className="notice-list">
        {thirdPartyNotices.map((notice) => (
          <li key={`${notice.name}@${notice.version}`}>
            <strong>{notice.name}</strong> {notice.version}
            <span>{notice.license}</span>
            {notice.copyright ? <small>{notice.copyright}</small> : null}
          </li>
        ))}
      </ul>
    </section>
  );
}

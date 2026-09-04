#!/usr/bin/env node
/**
 * Generate a ref code for a Messenger outreach link and print the full
 * shareable URL. Plain Node, no dependencies — matches the rest of this
 * repo (static site + one Vercel function, no build step).
 *
 * Usage:
 *   node scripts/gen-ref.js <source> [baseUrl] [hashPath]
 *
 * Examples:
 *   node scripts/gen-ref.js fb
 *   node scripts/gen-ref.js fb https://squeakycleanco.co
 *   SITE_URL=https://squeakycleanco.co node scripts/gen-ref.js ig
 *   node scripts/gen-ref.js referral https://squeakycleanco.co "#/"
 *
 * baseUrl defaults to the SITE_URL env var, then to a placeholder that
 * makes it obvious you still need to supply your real domain.
 *
 * IMPORTANT: ref has to land in the query string, which means BEFORE the
 * "#" — "https://domain/?ref=xyz#/apply", never "#/apply?ref=xyz". The
 * page reads location.search for ref, and everything after "#" is the
 * hash route, not the query string.
 */

'use strict';

const SLUG_CHARS = 'abcdefghijklmnopqrstuvwxyz0123456789';

function slug(len) {
  let out = '';
  for (let i = 0; i < len; i++) out += SLUG_CHARS[Math.floor(Math.random() * SLUG_CHARS.length)];
  return out;
}

function mmdd(d) {
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return mm + dd;
}

const source = (process.argv[2] || '').trim();
if (!source || !/^[a-z0-9]+$/i.test(source)) {
  console.error('Usage: node scripts/gen-ref.js <source> [baseUrl] [hashPath]');
  console.error('  <source>  short tag, letters/digits only — e.g. fb, ig, sms, referral');
  console.error('  baseUrl   defaults to $SITE_URL, e.g. https://squeakycleanco.co');
  console.error('  hashPath  defaults to "#/apply" (send people straight to the form)');
  process.exit(1);
}

const baseUrl = (process.argv[3] || process.env.SITE_URL || 'https://YOUR-DOMAIN-HERE').replace(/\/+$/, '');
const hashPath = process.argv[4] || '#/apply';

const ref = source.toLowerCase() + '-' + mmdd(new Date()) + '-' + slug(4);
const url = baseUrl + '/?ref=' + encodeURIComponent(ref) + hashPath;

console.log(ref);
console.log(url);

if (baseUrl.indexOf('YOUR-DOMAIN-HERE') > -1) {
  console.error('\nNo domain set — pass one as the 2nd argument or set SITE_URL, e.g.:');
  console.error('  node scripts/gen-ref.js ' + source + ' https://squeakycleanco.co');
}

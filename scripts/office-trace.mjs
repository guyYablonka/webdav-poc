#!/usr/bin/env node
// Summarises dav-service logs written with DAV_LOG_HEADERS=true into a readable
// request/response trace, to compare Office's real traffic with the README.
//
//   npm run dev 2>&1 | tee office.log      (then use Office)
//   npm run office-trace -- office.log
import { createReadStream } from 'node:fs';
import { createInterface } from 'node:readline';

const INTERESTING_REQUEST = [
  'user-agent',
  'depth',
  'if',
  'lock-token',
  'timeout',
  'content-type',
  'content-length',
  'if-match',
  'if-none-match',
  'if-modified-since',
  'translate',
  'overwrite',
  'destination',
];
const INTERESTING_RESPONSE = ['etag', 'lock-token', 'content-type', 'content-length', 'www-authenticate'];

const input = process.argv[2] ? createReadStream(process.argv[2]) : process.stdin;
const pending = new Map();

function pick(headers = {}, names) {
  return names
    .filter((n) => headers[n] !== undefined)
    .map((n) => `${n}: ${headers[n]}`)
    .join('\n      ');
}

for await (const line of createInterface({ input })) {
  const start = line.indexOf('{');
  if (start === -1) continue;
  let entry;
  try {
    entry = JSON.parse(line.slice(start));
  } catch {
    continue;
  }
  if (entry.service !== 'dav-service' || !entry.reqId) continue;
  if (entry.msg === 'dav request') pending.set(entry.reqId, entry);
  if (entry.msg !== 'dav response') continue;
  const req = pending.get(entry.reqId);
  pending.delete(entry.reqId);
  if (entry.url === '/healthz') continue;
  const time = new Date(entry.time).toISOString().slice(11, 23);
  console.log(`${time}  ${entry.method.padEnd(9)} ${entry.statusCode}  ${entry.url}  (${entry.ms} ms)`);
  const reqHeaders = pick(req?.headers, INTERESTING_REQUEST);
  if (reqHeaders) console.log(`   >  ${reqHeaders}`);
  const resHeaders = pick(entry.responseHeaders, INTERESTING_RESPONSE);
  if (resHeaders) console.log(`   <  ${resHeaders}`);
}

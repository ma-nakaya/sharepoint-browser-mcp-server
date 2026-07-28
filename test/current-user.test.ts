import assert from "node:assert/strict";
import test from "node:test";

import { parseCurrentUserResponse } from "../src/sharepoint/current-user.js";

test("parses an OData no-metadata current-user response", () => {
  assert.deepEqual(
    parseCurrentUserResponse(
      JSON.stringify({ Id: 42, Title: "Example User", IsSiteAdmin: false }),
    ),
    { id: 42, displayName: "Example User", isSiteAdmin: false },
  );
});

test("parses a verbose current-user response", () => {
  assert.deepEqual(
    parseCurrentUserResponse(
      JSON.stringify({ d: { Id: 7, Title: "Site Admin", IsSiteAdmin: true } }),
    ),
    { id: 7, displayName: "Site Admin", isSiteAdmin: true },
  );
});

test("drops email address and login name from the parsed result", () => {
  const result = parseCurrentUserResponse(
    JSON.stringify({
      Id: 42,
      Title: "Example User",
      IsSiteAdmin: false,
      Email: "user@example.com",
      LoginName: "i:0#.f|membership|user@example.com",
    }),
  );

  assert.deepEqual(Object.keys(result).sort(), ["displayName", "id", "isSiteAdmin"]);
});

test("rejects an incomplete current-user response", () => {
  assert.throws(
    () => parseCurrentUserResponse(JSON.stringify({ Id: 1, Title: "User" })),
    /IsSiteAdmin/u,
  );
});

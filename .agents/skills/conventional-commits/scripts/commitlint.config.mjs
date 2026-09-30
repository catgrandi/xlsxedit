const allowedTypes = [
  "build",
  "chore",
  "ci",
  "docs",
  "feat",
  "fix",
  "perf",
  "refactor",
  "revert",
  "style",
  "test",
];

const bundledFooterTokens = [
  "BREAKING CHANGE",
  "BREAKING-CHANGE",
  "Fixes",
  "Closes",
  "Resolves",
  "Refs",
  "See",
  "Co-Authored-By",
  "Reviewed-By",
  "Signed-Off-By",
  "Acked-By",
  "Reverts",
  "Deprecated",
];

// Tokens are matched case-insensitively, as the Conventional Commits
// specification requires, except the breaking-change tokens, which the
// specification requires to be uppercase.
const breakingTokenSet = new Set(["breaking change", "breaking-change"]);

// A trailer token is a word that may contain digits and hyphens, or the
// two-word "BREAKING CHANGE" form. The loose pattern tolerates space before
// the colon so near-miss trailers such as "Fixes : #123" are recognized and
// rejected as malformed instead of passing as prose.
const tokenPattern = /^[A-Za-z0-9][A-Za-z0-9-]*$/u;
const looseColonPattern = /^([A-Za-z0-9][A-Za-z0-9-]*(?: CHANGE)?)[\t ]*:/iu;
const strictColonPattern = /^([A-Za-z0-9][A-Za-z0-9-]*(?: CHANGE)?):[\t ]+\S/iu;
// The specification's second footer separator: "<token> #<reference>". Lines
// containing "://" are excluded so URLs with numeric fragments stay prose.
const hashSeparatorPattern = /^([A-Za-z0-9][A-Za-z0-9-]*)[\t ]+\S*#\d+$/u;

function splitBlocks(raw) {
  return raw
    .trimEnd()
    .split(/\r?\n[\t ]*\r?\n/u)
    .map((block) => block.split(/\r?\n/u));
}

function classifyLine(line) {
  const colonToken = line.match(looseColonPattern)?.[1] ?? null;
  const hashToken = line.includes("://") ? null : (line.match(hashSeparatorPattern)?.[1] ?? null);
  const token = colonToken ?? hashToken;

  return {
    hashForm: colonToken === null && hashToken !== null,
    line,
    token,
    trailerShaped: strictColonPattern.test(line) || (colonToken === null && hashToken !== null),
  };
}

function createFooterTokenRule(allowedFooterTokens) {
  const allowedTokenSet = new Set(allowedFooterTokens.map((token) => token.toLowerCase()));

  function checkTrailerLine(parsed) {
    if (parsed.hashForm) {
      return allowedTokenSet.has(parsed.token.toLowerCase())
        ? `write the reference footer as "${parsed.token}: <value>", not "${parsed.token} #..."`
        : `footer token "${parsed.token}" must match one of: ${allowedFooterTokens.join(", ")}`;
    }

    if (parsed.token === null || !strictColonPattern.test(parsed.line)) {
      return `footer line "${parsed.line}" must use <Token>: <value>; separate body prose from footers with a blank line`;
    }

    if (breakingTokenSet.has(parsed.token.toLowerCase())) {
      return parsed.token === parsed.token.toUpperCase()
        ? null
        : `breaking-change footer token must be uppercase: ${parsed.token.toUpperCase()}`;
    }

    return allowedTokenSet.has(parsed.token.toLowerCase())
      ? null
      : `footer token "${parsed.token}" must match one of: ${allowedFooterTokens.join(", ")}; ` +
          "reword a final paragraph that is not a footer so its lines do not parse as trailers";
  }

  return ({ raw }) => {
    const blocks = splitBlocks(raw);

    // Walk the trailing run of footer-shaped paragraphs. A paragraph is a
    // footer paragraph when it names an allowed token or every line is
    // trailer-shaped; anything else is body prose, which may freely contain
    // colons and URLs, and ends the walk.
    for (let index = blocks.length - 1; index >= 1; index -= 1) {
      const written = blocks[index].filter((line) => line.trim() !== "");

      // Extra blank lines are spacing rather than a paragraph, so they must
      // not end the run and hide a footer paragraph behind them.
      if (written.length === 0) {
        continue;
      }

      // A wholly indented paragraph is body content, such as a code block,
      // and it ends the run.
      const contentLines = written.filter((line) => !/^[\t ]/u.test(line));

      if (contentLines.length === 0) {
        break;
      }

      const parsedLines = contentLines.map(classifyLine);
      const isFooterBlock =
        parsedLines.some(
          (parsed) => parsed.token !== null && allowedTokenSet.has(parsed.token.toLowerCase()),
        ) || parsedLines.every((parsed) => parsed.trailerShaped);

      if (!isFooterBlock) {
        break;
      }

      let openTrailer = false;

      for (const parsed of parsedLines) {
        // A footer value may run onto the following lines until the next
        // token, so a line below a trailer continues it unless the line is
        // itself shaped like a trailer or names an allowed token with a
        // broken separator.
        const namesAllowedToken =
          parsed.token !== null && allowedTokenSet.has(parsed.token.toLowerCase());

        if (openTrailer && !parsed.trailerShaped && !namesAllowedToken) {
          continue;
        }

        const problem = checkTrailerLine(parsed);
        if (problem !== null) {
          return [false, problem];
        }

        openTrailer = true;
      }
    }

    return [true, "footer tokens are allowed"];
  };
}

const headerRules = {
  "header-max-length": ({ header }, _when, limit) => [
    [...header].length <= limit,
    `header must be ${limit} characters or fewer`,
  ],
  "header-target-length": ({ header }, _when, target) => [
    [...header].length <= target,
    `header should be ${target} characters or fewer`,
  ],
  "scope-format": ({ header, scope }) => {
    if (/^\w*\(\)/u.test(header ?? "")) {
      return [false, "scope parentheses must not be empty; omit the scope instead"];
    }

    return [
      scope == null || scope === "" || /^[a-z0-9]+(?:[-_][a-z0-9]+)*$/u.test(scope),
      "scope must name a module or component in lowercase letters and digits, joined by " +
        "single '-' or '_' separators, and must not name a file, extension, or path",
    ];
  },
  "subject-lowercase-start": ({ subject }) => [
    typeof subject === "string" && !/^\p{Lu}/u.test(subject),
    "subject must not start with an uppercase letter",
  ],
};

// A repository may admit trailers that its tooling writes, such as the
// session link an authoring tool stamps on a commit, without editing this
// file: import `createConfig` from a configuration of its own and pass them
// as `footerTokens`. Each must have the shape of a trailer token, so that a
// misconfigured extension fails when the configuration loads rather than
// admitting nothing in silence.
export function createConfig({ footerTokens = [] } = {}) {
  const allowedFooterTokens = [...bundledFooterTokens];

  for (const token of footerTokens) {
    if (typeof token !== "string" || !tokenPattern.test(token)) {
      throw new Error(
        `footer token ${JSON.stringify(token)} must be one word of letters, digits, and hyphens`,
      );
    }

    if (!allowedFooterTokens.some((allowed) => allowed.toLowerCase() === token.toLowerCase())) {
      allowedFooterTokens.push(token);
    }
  }

  return {
    defaultIgnores: false,
    parserPreset: {
      name: "conventional-commits-policy",
      parserOpts: {
        headerCorrespondence: ["type", "scope", "subject"],
        headerPattern: /^(\w*)(?:\(([^)]*)\))?!?: (.+)$/u,
      },
    },
    plugins: [
      {
        rules: {
          "footer-token-enum": createFooterTokenRule(allowedFooterTokens),
          ...headerRules,
        },
      },
    ],
    rules: {
      "body-leading-blank": [2, "always"],
      "breaking-change-exclamation-mark": [2, "always"],
      "footer-leading-blank": [2, "always"],
      "footer-token-enum": [2, "always"],
      "header-max-length": [2, "always", 72],
      "header-target-length": [1, "always", 50],
      "header-trim": [2, "always"],
      "scope-case": [2, "always", "lower-case"],
      "scope-format": [2, "always"],
      "subject-empty": [2, "never"],
      "subject-full-stop": [2, "never", "."],
      "subject-lowercase-start": [2, "always"],
      "type-case": [2, "always", "lower-case"],
      "type-empty": [2, "never"],
      "type-enum": [2, "always", allowedTypes],
    },
  };
}

export default createConfig();

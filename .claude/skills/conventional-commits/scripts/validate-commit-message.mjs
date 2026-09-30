import { spawnSync } from "node:child_process";
import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const startDirectory = process.cwd();
const configPath = fileURLToPath(new URL("./commitlint.config.mjs", import.meta.url));

function findRepositoryRoot(directory) {
  let current = path.resolve(directory);

  while (true) {
    if (fs.existsSync(path.join(current, ".git"))) {
      return current;
    }

    const parent = path.dirname(current);
    if (parent === current) {
      return null;
    }

    current = parent;
  }
}

function realPath(target) {
  try {
    return fs.realpathSync(target);
  } catch {
    return path.resolve(target);
  }
}

function isInside(root, target) {
  const relative = path.relative(realPath(root), realPath(target));
  return relative !== "" && !relative.startsWith("..") && !path.isAbsolute(relative);
}

// Node resolution walks up past the repository. Anchoring at the repository
// root and rejecting a result outside it keeps a validation run from silently
// using an unrelated checkout's commitlint, or none at all.
const repositoryRoot = findRepositoryRoot(startDirectory) ?? startDirectory;
const requireFromRepository = createRequire(path.join(repositoryRoot, "package.json"));

let commitlintCli;
try {
  commitlintCli = requireFromRepository.resolve("@commitlint/cli/cli.js");
} catch {
  commitlintCli = null;
}

if (commitlintCli === null || !isInside(repositoryRoot, commitlintCli)) {
  const reason =
    commitlintCli === null
      ? "Cannot find @commitlint/cli"
      : `Found @commitlint/cli outside the repository at ${commitlintCli}, which may not match this repository's policy or lockfile`;

  console.error(
    `${reason}. Looked in ${repositoryRoot}. Use the repository's own validator, review the message manually, or install commitlint only with the user's approval.`,
  );
  process.exit(2);
}

// Commitlint's parser degrades sharply on a single very long line: a 2000
// character line takes seconds and a 4000 character line does not finish.
// Reporting the line beats hanging with no output.
const parseableLineLength = 2000;
const commitlintArguments = process.argv.slice(2);
const readsStandardInput = commitlintArguments.length === 0 && !process.stdin.isTTY;
let message = null;

if (readsStandardInput) {
  message = fs.readFileSync(0, "utf8");

  const overlongLine = message
    .split(/\r?\n/u)
    .findIndex((line) => line.length > parseableLineLength);

  if (overlongLine !== -1) {
    console.error(
      `Line ${overlongLine + 1} of the message is longer than ${parseableLineLength} characters, which commitlint cannot parse in reasonable time. Wrap the body and validate again.`,
    );
    process.exit(2);
  }
}

const result = spawnSync(
  process.execPath,
  [commitlintCli, "--config", configPath, ...commitlintArguments],
  {
    cwd: startDirectory,
    env: { ...process.env, NO_COLOR: "1" },
    ...(message === null
      ? { stdio: "inherit" }
      : { input: message, stdio: ["pipe", "inherit", "inherit"] }),
  },
);

if (result.error) {
  console.error(`Unable to run commitlint: ${result.error.message}`);
  process.exit(2);
}

process.exit(result.status ?? 2);

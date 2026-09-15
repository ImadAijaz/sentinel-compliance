#!/usr/bin/env bash
# Make ONE git push update BOTH deployments.
#
# THE PROBLEM THIS SOLVES
# This project is deployed twice, from two SEPARATE GitHub repos:
#   iticcotx/sentinel-compliance   -> https://sentinel-compliance-delta.vercel.app
#   ImadAijaz/sentinel-compliance  -> https://sentinel-compliance-kappa.vercel.app
# Nothing connects the two repos. Pushing to one leaves the other behind, and the
# two live sites then show DIFFERENT compliance answers — which for a compliance
# dashboard is worse than being down. It happened: delta reached cloud122 while
# kappa sat at cloud116, six versions and twenty commits behind, for days.
#
# THE FIX
# Give each remote two push URLs. `git push` then writes to both repos every
# time, and the two sites cannot drift apart again. Remote config lives in
# .git/config, which is NOT committed — so run this once after any fresh clone.
#
#   bash tools/setup-remotes.sh
#
# A better long-term fix is to point BOTH Vercel projects at the SAME repo (or
# retire one of the two deployments entirely). That needs the Vercel dashboard.
# Until then, this keeps them in step.

set -e
KAPPA="https://github.com/ImadAijaz/sentinel-compliance.git"
ORIGIN="https://github.com/iticcotx/sentinel-compliance.git"

git remote get-url origin >/dev/null 2>&1 || git remote add origin "$ORIGIN"
git remote get-url kappa  >/dev/null 2>&1 || git remote add kappa  "$KAPPA"

# The first --push REPLACES the default push URL; --add appends the second.
git remote set-url --push       origin "$ORIGIN"
git remote set-url --add --push origin "$KAPPA"
git remote set-url --push       kappa  "$KAPPA"
git remote set-url --add --push kappa  "$ORIGIN"

echo "Done. Either of these now updates BOTH sites:"
echo "    git push origin main:main"
echo "    git push kappa  main:main"
echo
git remote -v
echo
echo "Verify after pushing — both must report the same version:"
echo "    curl -s https://sentinel-compliance-delta.vercel.app/ | grep -o 'cloud[0-9]*' | sort -u"
echo "    curl -s https://sentinel-compliance-kappa.vercel.app/ | grep -o 'cloud[0-9]*' | sort -u"

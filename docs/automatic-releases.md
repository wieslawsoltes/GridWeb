# Verified main releases

A successful CI run triggered by a push to this repository's main branch starts Publish verified main version. It checks that the qualified SHA is still the current main head, validates the package version, and creates an immutable version tag. Existing tags are never moved. A main commit reusing an already released version is skipped; increment package.json and package-lock.json for the next release.

The job then calls the same Release and npm workflow used by explicit tags and manual retries, passing the exact verified SHA. Release/browser/package checks rerun on the tag before the exact tarball is published with NPM_TOKEN and provenance. Public registry integrity is checked before the draft GitHub release is published. The privileged workflow does not execute fork pull-request code or consume pull-request artifacts.

GitHub Pages deploys independently after testing the main source. Publication/deployment status must be checked in Actions; workflow source alone is not a successful release.

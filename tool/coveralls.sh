#!/bin/bash

# Fast fail the script on failures.
set -e

# Install dart_coveralls
# Gather and send coverage data.
if [ "$COVERALLS_TOKEN" ] && [ "$TRAVIS_DART_VERSION" = "stable" ]; then
  pub global activate dart_coveralls
  pub global run dart_coveralls report \
    --exclude-test-files \
    test/io_test.dart
fi

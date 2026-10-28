"""Positive and negative fixtures for committed content checks."""

import unittest

import yaml

from check_content import UniqueKeysLoader, markdown_errors


class ContentTests(unittest.TestCase):
    def test_markdown_links_and_fences(self):
        files = {"docs/guide.md", "docs/target.md"}
        self.assertEqual(markdown_errors("docs/guide.md", "[ok](target.md)\n```sh\n[ignore](missing.md)\n```\n", files), [])
        self.assertEqual(len(markdown_errors("docs/guide.md", "[bad](missing.md)\n```\n", files)), 2)

    def test_yaml_valid_and_duplicate(self):
        self.assertEqual(yaml.load("name: check\nsteps:\n  - run: echo ok\n", Loader=UniqueKeysLoader)["name"], "check")
        with self.assertRaises(yaml.YAMLError):
            yaml.load("name: one\nname: two\n", Loader=UniqueKeysLoader)
        with self.assertRaises(yaml.YAMLError):
            yaml.load("name: [unterminated\n", Loader=UniqueKeysLoader)


if __name__ == "__main__":
    unittest.main()

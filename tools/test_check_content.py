"""Positive and negative fixtures for committed content checks."""

import unittest

import yaml
from check_content import UniqueKeysLoader, markdown_errors


class ContentTests(unittest.TestCase):
    def test_markdown_links_and_fences(self) -> None:
        files = {"docs/guide.md", "docs/target.md"}
        self.assertEqual(markdown_errors("docs/guide.md", "[ok](target.md)\n```sh\n[ignore](missing.md)\n```\n", files), [])
        self.assertEqual(len(markdown_errors("docs/guide.md", "[bad](missing.md)\n```\n", files)), 2)

    def test_fragments_images_reference_and_html_links(self) -> None:
        docs = {"README.md": "<a id=\"-status\"></a>\n<a id=\"-status-1\"></a>\n<a id=\"stable\"></a>\n",
                "docs/guide.md": ""}
        names = set(docs) | {"image.png"}
        good = "[ok](../README.md#-status) [duplicate](../README.md#-status-1)\n" \
               "[stable][ref]\n\n[ref]: ../README.md#stable\n" \
               "![image](../image.png) <a href=\"../README.md#stable\">ok</a>"
        self.assertEqual(markdown_errors("docs/guide.md", good, names, docs), [])
        for bad in ("[bad](../README.md#status)", "![bad](missing.png)",
                    '<a href="../README.md#missing">bad</a>',
                    "[bad][ref]\n\n[ref]: missing.md", "[bad](../../outside.md)"):
            with self.subTest(bad=bad):
                self.assertTrue(markdown_errors("docs/guide.md", bad, names, docs))

    def test_yaml_valid_and_duplicate(self) -> None:
        self.assertEqual(yaml.load("name: check\nsteps:\n  - run: echo ok\n", Loader=UniqueKeysLoader)["name"], "check")
        with self.assertRaises(yaml.YAMLError):
            yaml.load("name: one\nname: two\n", Loader=UniqueKeysLoader)
        with self.assertRaises(yaml.YAMLError):
            yaml.load("name: [unterminated\n", Loader=UniqueKeysLoader)


if __name__ == "__main__":
    unittest.main()

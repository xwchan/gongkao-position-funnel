"""Conservative, text-only graduate-status hints extracted from free-form notes."""
import re

YEAR = r"2027(?:年)?届(?:应届)?(?:高校)?毕业生|2027年应届(?:高校)?毕业生|2027应届(?:高校)?毕业生"
FRESH = r"应届(?:高校)?毕业生"
PREFIX = (r"(?:仅限|只限|限招|仅招(?:录|收)?|(?<!不)限|只招收|招考对象为|招录对象为|"
          r"招收对象为|招考对象是|招录对象是|面向|仅面向|招聘对象为|本职位招录|本职位招考)\s*[“\"\u201c\u201d]?")
BARE = r"(?:^|[。；;\n])\s*(?:\d+[.．、]\s*)?{}\s*(?:[,，；;。\n（(]|$)"
YEAR_EXPLICIT = re.compile(PREFIX + "(?:" + YEAR + ")")
FRESH_EXPLICIT = re.compile(PREFIX + "(?:" + FRESH + ")")
YEAR_BARE = re.compile(BARE.format("(?:" + YEAR + ")"))
FRESH_BARE = re.compile(BARE.format("(?:" + FRESH + ")"))
NON_EXPLICIT = re.compile(PREFIX + r"(?:非应届(?:高校)?毕业生|非2027届(?:高校)?毕业生)")


def classify(note: str) -> str:
    """Return 2027/fresh/non/unspecified; 'unspecified' is not an eligibility verdict."""
    if NON_EXPLICIT.search(note):
        return "non"
    if YEAR_EXPLICIT.search(note) or YEAR_BARE.search(note):
        return "2027"
    if FRESH_EXPLICIT.search(note) or FRESH_BARE.search(note):
        return "fresh"
    return "unspecified"

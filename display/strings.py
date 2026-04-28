# ui/strings.py
# All user-facing text constants for GrayWolfe.
# Ported from old_vba/mod4b_DefineThings.bas (DefineFormDefaults + DefinePopupText).
# No imports. No logic beyond string formatting.

# ---------------------------------------------------------------------------
# Static constants
# ---------------------------------------------------------------------------

SEARCH_PLACEHOLDER = (
    "Search one or multiple I selectors. Input / paste as many as you want HERE."
)

ADD_DEFAULT_PLACEHOLDER = (
    "Input / Paste I selectors to add HERE. GrayWolfe auto detects data types. "
    "ONE RECORD SET PER LINE (see examples below).\n\n"
    "Data can be delimited (separated) by comma, semicolon, pipe, tab, or anything you can imagine.\n\n"
    "Input example (pipe-delimited):\n"
    "Kim Jong Un | Residence No. 55 National Hwy 65 Pyongyang, DPRK | linkedin.com/in/superstar1984|@michaeljordanlover84\n"
    "Kim Ju Ae | Dorm 1A Wellington Square Oxford, UK | linkedin.com/in/AI-Dev | @zoomerPrincess2012\n\n"
    "Another example (comma-delimited):\n"
    "entrepreneur1950@gmail.com, 202-123-4567, Bob Smith\n"
    "faciliatorForForeignPower@hotmail.com, 1-701-987-6543, Tim Jones\n"
    "literallyBuyingNukes@yahoo.com, 3025439876, Susie Johnson\n\n"
    "Another example (semicolons):\n"
    "Kim Jong Il; nerd.loser@gmail.com; linkedin.com/in/source-of-the-problem007\n"
    "Robert Robertson; [leave blank]; linkedin.com/in/i-dont-want-more-examples"
)

ADD_UNRELATED_PLACEHOLDER = (
    "Input / Paste UNRELATED I selectors to add HERE. "
    "For bulk lists where connections are unknown or nonexistent.\n\n"
    "Input example:\n"
    "Kim Il Sung; CCP_1931@email.org; Robertson Roberts; 123 Fake St. Wilmington, DE 19809; etc.\n\n"
    "[TLDR: This import stores each selector separately. "
    "If multiple selectors are owned by the same dude please use the OTHER import.]"
)

TOKEN_PLACEHOLDER = "[Paste YOUR S API Token HERE]"

DISCLAIMER_SEARCH = (
    "DISCLAIMER: GrayWolfe is NOT an F data repository and does NOT have unique I data. "
    "GrayWolfe queries / displays data from the F's official data repository."
)

DISCLAIMER_ADD = (
    "DISCLAIMER: Only add selectors ALREADY IN S; "
    "any selector not ALREADY IN S is auto deleted and removed from GrayWolfe."
)

CONFIRM_RESET_TARGETS = (
    "Are you sure you want to reset the target data?!?\n\n"
    "This deletes Target edits since your last save and CANNOT be undone.\n\n"
    "Click Yes to proceed, No to cancel."
)

WARN_CHANGE_TARGET_ID = (
    "Are you sure you want to change the targetId?!?!1!\n\n"
    "Hit Yes to proceed, NO to cancel (pls hit NO).\n\n"
    "[TLDR: This is a very bad idea, and will probably break everything, "
    "I strongly recommend against it.]\n\n"
    "(But I also believe in freedom, so I will let you do it, "
    "but please dont unless you really know what you're doing)."
)

# ---------------------------------------------------------------------------
# Dynamic builder functions
# ---------------------------------------------------------------------------

_VOWELS = ("a", "e", "i", "o", "u")


def type_wrong_text(selector: str, sel_type: str) -> str:
    """VBA typeWrong — confirmation when detected type doesn't match user's choice."""
    first_char = sel_type.strip()[:1].lower()
    article = "an" if first_char in _VOWELS else "a"
    return (
        f'Are you sure "{selector}" is {article} {sel_type}??\n\n'
        "Click YES to override, click NO to cancel and resubmit"
    )


def type_null_text(selector: str) -> str:
    """VBA typeNull — when type detection completely fails."""
    return (
        f'GrayWolfe FAILED to detect a type for "{selector}"\n\n'
        "GrayWolfe's *detection algorithm* [a billion if statements and regexs] "
        "cant figure it out (#sad_emoticon).\n\n"
        f'Do you want to skip "{selector}" and keep going?\n\n'
        "Click Yes to skip, click No to set the type with the drop down / try again."
    )


def import_result_text(submitted: int, inserted: int, skipped: int) -> str:
    """Unified searchDisplayDefault/Unrelated — shown after import completes.

    submitted: total unique selectors submitted
    inserted:  how many were new and uploaded
    skipped:   how many already existed in GrayWolfe
    """
    # Header: X unique selector(s) submitted
    if submitted == 1:
        text = "1 unique selector submitted."
    else:
        text = f"{submitted} unique selectors submitted."
    text += "\n\n"

    # All existing — none new
    if inserted == 0:
        if skipped == 1:
            text += "Selector searched is already in GrayWolfe!"
        else:
            text += f"All {skipped} selectors are already in GrayWolfe!"
        text += "\n\nTo see full search results, click OK."
        return text

    # Some or all new
    if inserted == 1:
        # inserted==submitted: only 1 selector total and it's new
        # inserted!=submitted: 1 new among multiple (the rest were skipped)
        if inserted == submitted:
            text += "Searched selector is NEW and was successfully uploaded to GrayWolfe!"
        else:
            text += "1 selector is NEW and was successfully uploaded to GrayWolfe!"
    else:
        if inserted == submitted:
            text += f"All {inserted} selectors are NEW and were successfully uploaded to GrayWolfe!"
        else:
            text += f"{inserted} selectors are NEW and were successfully uploaded to GrayWolfe!"

    text += "\n\n"

    # No existing at all — intentionally no "click OK" prompt here (mirrors VBA)
    if skipped == 0:
        return text

    # Mixed: some already existed
    if skipped == 1:
        text += "1 selector is already in GrayWolfe."
    else:
        text += f"{skipped} selectors are already in GrayWolfe."

    text += "\n\nTo see the full search results, click OK."
    return text


def rate_limit_text(selector: str, sel_idx: int, total_selectors: int) -> str:
    """VBA rateLimitText — shown in status bar during rate-limit pauses.

    sel_idx: 0-based index of the current selector.
    total_selectors: actual total count of selectors being searched.
    Returns a single-line string suitable for the status bar.
    """
    current = sel_idx + 1
    total = total_selectors
    return (
        f'Searching "{selector}" ({current}/{total}) — '
        "taking a 10s rate-limit break so we don't DDOS S (inshaAllah)... "
        "spin in your chair 5x whistling Yankee Doodle and we'll resume."
    )


def search_check_text(selector: str, num_found: int) -> str:
    """VBA searchCheck — mid-search confirmation when S returns many hits."""
    selector = selector.strip()
    return (
        f'Your search of "{selector}" produced {num_found} hits!!!\n\n'
        "Do you want to wait for this search to finish (it will prob take a while), "
        "or do you want the tool to SKIP this item, so everything else finishes faster?\n\n"
        f'Click YES to continue searching "{selector}", NO to SKIP it.'
    )

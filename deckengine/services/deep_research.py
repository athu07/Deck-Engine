# -*- coding: utf-8 -*-
"""
deep_research.py  --  live web research for a client account, ADDITIVE to
(never replacing) the existing "upload a research/profile file" flow (owner's
spec, 2026-07-08: keep the manual upload option, add auto-research alongside
it).

The ONLY place in this app that makes a live internet call beyond OpenAI's own
text-completion endpoints -- every other AI call here is closed-book (reads
only what the user pasted/uploaded). Uses OpenAI's own Responses API
`web_search` tool, so this needs no new vendor or API key beyond the
OPENAI_API_KEY already configured for matching/generation.

strategic_brief() is the owner's own consulting-grade research prompt
(2026-07-08, given verbatim -- adapted here only to take client_name/
stakeholder_name as real variables instead of a fixed example, and with one
addition: an explicit grounding instruction, since a "cite a source" ROI
figure that's actually invented is a much worse failure mode for a strategic
brief than for a quick company blurb -- this goes in front of a real
executive). This REPLACES the earlier two separate, shorter research_company()/
research_stakeholder() calls -- the owner's prompt is inherently a single
joint deliverable ("a brief to present to X at Y"), not two independent ones.

Fail-safe like every other AI call in this app: any error (network, quota,
tool unavailable, nothing found) returns "" rather than raising, so a failed
auto-research never blocks a build -- the salesperson just falls back to
typing notes or uploading a brief, same as if they'd never clicked the button.

The result is handed back to the SALESPERSON to review/edit before it's used
for matching (see /research_account in web/api.py) -- never silently baked
into a build, matching the human-review-gate design already used for every
other AI output in this app (e.g. the Create-with-AI case-study flow).
"""

from deckengine.services.infra import load_env

MODEL = "gpt-4o-mini"


def _client():
    load_env()
    from openai import OpenAI
    return OpenAI()


def _search(prompt):
    try:
        resp = _client().responses.create(
            model=MODEL, tools=[{"type": "web_search"}], input=prompt)
        return (resp.output_text or "").strip()
    except Exception:
        return ""


def strategic_brief(company_name, stakeholder_name="", industry="",
                    profile_text="", research_text=""):
    """An executive-grade strategic account brief -- consulting insight and
    innovation-opportunity discovery, not generic company research. Owner's
    own prompt structure (2026-07-08), used close to verbatim. Requires a
    company name; the stakeholder is optional (the brief adapts if there's no
    named recipient yet). profile_text/research_text -- owner's spec,
    2026-07-09: if the salesperson already attached a stakeholder profile or
    research file on the SAME form, ground this live research in that real
    content first (PRIMARY source), THEN go beyond it via live search --
    instead of researching purely from a typed name with zero awareness of
    what's already on file for this account. Empty string if nothing found or
    the search fails."""
    company_name = (company_name or "").strip()
    if not company_name:
        return ""
    stakeholder_name = (stakeholder_name or "").strip()
    profile_text = (profile_text or "").strip()
    research_text = (research_text or "").strip()
    who = f"{stakeholder_name} at {company_name}" if stakeholder_name else company_name
    persona = stakeholder_name or "the buying team"
    attached = ""
    if profile_text or research_text:
        parts = []
        if profile_text:
            parts.append("STAKEHOLDER PROFILE (their own real background -- a stakeholder's "
                         "own certifications and career-long expertise are a strong signal "
                         "of what they'd care about, even under a newer, more senior title; "
                         "ground the brief in this, don't contradict it):\n\"\"\"\n"
                         + profile_text[:6000] + "\n\"\"\"")
        if research_text:
            parts.append("ALREADY-ATTACHED RESEARCH/NOTES FOR THIS ACCOUNT:\n\"\"\"\n"
                         + research_text[:6000] + "\n\"\"\"")
        attached = ("ATTACHED CONTEXT (already on file for this account -- treat as your "
                    "PRIMARY, most reliable source for who this stakeholder actually is and "
                    "what this account is dealing with):\n\n" + "\n\n".join(parts) + "\n\n")
    intersection = ""
    if stakeholder_name:
        intersection = (
            f"CRITICAL -- this brief must sit at the INTERSECTION of two things, not just "
            f"one: (1) what {company_name} genuinely needs (its strategic priorities, "
            f"gaps, transformation initiatives), AND (2) what {stakeholder_name} "
            f"PERSONALLY would recognise, understand, and get genuinely excited about, "
            f"based on their own real background, expertise, and current role. A "
            f"technically-sound recommendation that's personally irrelevant to "
            f"{stakeholder_name} will not move them to champion it internally -- they "
            f"have to see THEMSELVES in it, not just their employer. This is a brief FOR "
            f"a specific person, not generic research ABOUT a company that happens to "
            f"employ them.\n\n"
            f"Every accelerator, gap, and opportunity you propose must be filtered "
            f"through BOTH lenses at once: does {company_name} need this, AND would "
            f"{stakeholder_name} personally resonate with it and be equipped to champion "
            f"it, given what they've actually done in their career? If a stakeholder "
            f"profile is attached above, deliberately favour accelerators and technology "
            f"areas that connect to their own real, substantive expertise over generic "
            f"company-wide options that don't touch their specific domain. Naming the "
            f"CATEGORY of their expertise (e.g. \"SAP integration\") is not enough -- at "
            f"least one accelerator must name the SPECIFIC platform, product, or "
            f"certification actually in their profile (e.g. if they hold an SAP BTP or "
            f"ABAP-for-HANA certification, an accelerator built around THAT, not a vague "
            f"\"SAP\" mention) so it reads as built for them specifically, not for anyone "
            f"with a similar job title -- that "
            f"intersection, not company research alone, is what actually wins a "
            f"project.\n\n"
        )
    prompt = (
        f'Act as a senior strategy consultant. Research and write an executive-grade '
        f'strategic brief to present to {who}'
        + (f" ({industry})" if industry else "") + ' to identify high-value synergies, '
        'transformation opportunities, capability alignments, and solution areas that '
        'we can potentially pitch.\n\n'
        + attached
        + intersection +
        'Research Expectations: Do not restrict the analysis to existing knowledge-base '
        'documents or attached files. Go significantly beyond that and research broader '
        'industry trends, company priorities, technology landscape, leadership focus '
        'areas, and strategic initiatives to uncover meaningful synergies and executable '
        'opportunities.\n\n'
        'Identify:\n'
        '- Current transformation priorities and strategic initiatives\n'
        '- Potential pain points, gaps, inefficiencies, or modernization opportunities. '
        'Named incumbent platforms -- their capabilities AND specific gaps\n'
        '- Three named, buildable accelerators that fill those gaps -- each chosen so at '
        'least one, ideally more, directly connects to the stakeholder\'s OWN real '
        'expertise/background where one is known (see the intersection rule above), not '
        'just the company\'s needs in the abstract -- each with: problem solved, core '
        'capabilities, platform-vs-content-layer distinction, augmentation framing, '
        '60-90 day pilot design, ROI/value case, and incumbent gap filled\n'
        '- Potential use-case synergies in a few lines each: the problem statement, the '
        'solution, and industry-standard quantified results\n'
        '- Technology, platform, AI, automation, or operational synergies\n'
        '- Opportunities where our capabilities, accelerators, frameworks, or services '
        'can add value\n'
        '- Innovative solution concepts, PoCs, quick wins, and long-term transformation '
        'themes\n'
        '- Executive-level narratives and strategic conversations that would resonate '
        'with this stakeholder specifically -- because the opportunity genuinely touches '
        'what they personally know and care about, not just because the tone is '
        'addressed to them\n\n'
        f"Frame everything in {persona}'s language. Position the accelerators as "
        'augmenting -- never replacing -- their existing stack and function. Every ROI '
        'figure must cite a source.\n\n'
        'The output should feel like strategic account intelligence combined with '
        'consulting insight and innovation opportunity discovery -- not generic company '
        'research.\n\n'
        'Ground every claim in what you actually find. Where you genuinely cannot find '
        'a real source for a specific figure or platform name, say so plainly instead of '
        'inventing one -- an unearned ROI number or a fabricated incumbent name in front '
        'of a real executive is a worse outcome than an honest gap.'
    )
    return _search(prompt)

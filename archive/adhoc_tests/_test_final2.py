import os, skills, assembler

ctx = {
    "work_types": ["WORKFORCE"],
    "industry": "BANKING",
    "transcript": "This is an RFI. We need Java and QA automation testing. Client is SocGen.",
    "client_name": "SocGen"
}

os.makedirs("output", exist_ok=True)
out = "output/_TEST_branded_v2.pptx"
assembler.build_deck(["CS01", "CS02", "CS07", "CS08"], out=out)

cands = skills.candidates(ctx)
print("Candidates:", [c["id"] for c in cands])

order = ["CS01", "CS02"] + [c["id"] for c in cands] + ["CS07", "CS08"]
n = skills.build_into(out, order, cands)
print(f"Built {n} skills slides -> {out}")
print("Open output/_TEST_branded_v2.pptx to check the design.")

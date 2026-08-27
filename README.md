# Uplift Modeling Demo — Credit Card Prescreen Targeting (Synthetic Data)

A hand-rolled **X-learner** uplift model, built with synthetic data to demonstrate the same methodology as a real credit-card prescreen targeting project. The real project's data and code are proprietary to the bank it was built at — this notebook reproduces the technique end-to-end so it can be reviewed, run, and discussed openly.

## What's here

- `chase_uplift_demo.ipynb` — the full notebook: simulate a randomized test/control campaign with a known heterogeneous treatment effect, fit an X-learner on scikit-learn (no black-box causal-ML package), validate with a Qini curve *and* a direct check against ground truth (only possible because the data is synthetic), and translate the model into a targeting decision.
- `chase_uplift_demo.html` — a static rendered view of the notebook (no Jupyter install needed).
- `build_notebook.py` — the script that generates the notebook. Re-run it to regenerate `chase_uplift_demo.ipynb` from scratch, or edit the data-generating process / model to experiment.
- `qini_curve.png`, `tau_scatter.png` — the two key validation plots, standalone.
- `uplift_model_deck.pptx` — a 16-slide walkthrough deck: why uplift over a response model, dedicated slides on S-Learner and T-Learner (how each trains and scores in the credit-card prescreen example, pros/cons, and why not chosen), the comparison table and why X-Learner wins, a full numeric walkthrough of the X-learner's four layers (outcome models → pseudo effects → effect models → weighted fusion), the Qini validation results above, and a one-minute summary (with a bilingual script in the speaker notes of the last slide).
- `build_deck.js` — the `pptxgenjs` script that generates the deck. Re-run with `node build_deck.js` after editing to regenerate `uplift_model_deck.pptx`.

## Why an uplift model instead of a response model

A plain response model predicts who's likely to respond to an offer — but that includes people who'd respond anyway ("sure things"), so targeting them wastes budget. An uplift model predicts the *incremental* effect of the offer per person, so targeting can focus on **persuadables**: the people whose decision the offer actually changes.

## Key result

In the simulated population here, the overall (average) incremental response from mailing everyone is close to flat — gains among persuadables are largely offset by a segment where the offer backfires ("sleeping dogs"). But the top uplift decile shows a clearly positive incremental lift. The aggregate number alone would suggest the campaign doesn't work; the uplift model reveals the real, exploitable heterogeneity hiding underneath it.

The X-learner's estimated uplift score correlates with the true (simulated) treatment effect at Spearman ρ ≈ 0.93, and captures roughly 96% of the oracle's achievable Qini gain.

## Run it

```bash
pip install numpy pandas scikit-learn matplotlib scipy jupyter
jupyter notebook chase_uplift_demo.ipynb
```

Or regenerate from scratch:

```bash
python build_notebook.py
jupyter nbconvert --to notebook --execute chase_uplift_demo.ipynb --output chase_uplift_demo.ipynb
```

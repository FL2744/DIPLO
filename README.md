# DIPLO: Public-Post Search, Translation, and Mapping

`DIPLO.py` is an interactive research script that searches public Nitter pages,
parses visible posts, optionally translates search terms and posts, infers broad
public locations, and creates tabular and map outputs.

The script does not bypass login walls, CAPTCHAs, private accounts, access
controls, anti-bot systems, or rate limits. Nitter instances can be unavailable
or change behavior, so live collection may sometimes fail independently of this
script. A `saved_html` mode is available for parsing pages you saved manually.

## What it creates

Depending on the settings in `DIPLO.py`, a run can produce:

- A CSV file containing collected and enriched post data
- A formatted Excel workbook
- An interactive HTML map
- A local geocoding cache

Inferred locations are broad, model-generated estimates. They are not verified
geotags and should not be treated as precise personal locations.

## First-time setup

Open Terminal, change to your project directory, and create a virtual Python
environment:

```bash
cd "your project directory"
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
python3 DIPLO.py
```

The script checks for its Python packages and installs missing ones into the
active virtual environment. Internet access is required for package installation,
LLM requests, live page retrieval, geocoding, and map tiles.

## Normal startup after the first run

From your project directory:

```bash
source .venv/bin/activate
python3 DIPLO.py
```

When finished, leave the virtual environment with:

```bash
deactivate
```

## Choose an LLM provider

When translation or location inference is enabled, the script asks you to
choose one of two providers.

### OpenAI API

The OpenAI menu offers:

1. `gpt-5.6-luna` — default; intended for cost-sensitive, high-volume work
2. `gpt-5.6-terra` — balances capability and cost
3. `gpt-5.6-sol` — flagship capability
4. A different OpenAI model ID entered by the user

Enter an OpenAI Platform API key when prompted.

### Virginia Tech ARC LLM API

The ARC menu offers:

1. `gpt-oss-120b`
2. `DeepSeek-V4-Flash`
3. `GLM-5.2`
4. `Kimi-K3`
5. A different ARC model ID entered by the user

The script uses ARC's OpenAI-compatible endpoint at
`https://llm-api.arc.vt.edu/api/v1`. Virginia Tech students, faculty, and staff
can create a personal API key at <https://llm.arc.vt.edu/> under **User profile
> Settings > Account > API keys**.

For both providers, key input is hidden. The key is retained only in memory for
the current run and is not written into the script or output files. Never share
an API key.

## Configure a run

The main settings are near the top of `DIPLO.py`:

- `MODE`: use `"live"` to query Nitter or `"saved_html"` to parse saved pages
- `NITTER_BASE_URLS`: Nitter instances used in live mode
- `SAVED_HTML_FILES`: input pages used in saved-HTML mode
- `TARGET_LANGUAGE`: translation target
- `SINCE_DATE` and `UNTIL_DATE`: optional date bounds
- `MAX_POSTS_PER_QUERY` and `MAX_PAGES_PER_QUERY`: collection limits
- `TRANSLATE_POSTS`: enable or disable post translation
- `INFER_LOCATIONS`: enable or disable model-based location inference
- `CREATE_MAP`: enable or disable the interactive map
- `DEFAULT_SEARCH_TERMS`: terms offered at startup

Start with small post and page limits while confirming that the selected Nitter
instance and provider work correctly.

## During a run

The script asks for search terms and optional languages into which those terms
should be translated. Original terms are always retained. Translated variants
are added as extra searches.

API requests and public geocoding can take time. Keep Terminal open until the
script reports the output filenames.

## Troubleshooting

### `externally-managed-environment`

Activate the project virtual environment before running the script:

```bash
source .venv/bin/activate
python3 DIPLO.py
```

### API authentication error

Run the script again and enter a valid key for the provider you selected. An
OpenAI key and an ARC key are not interchangeable.

### ARC access

ARC's shared API is intended for eligible Virginia Tech users. Generate the key
through your own ARC web profile and keep it confidential.

### No live Nitter results

The configured public instance may be unavailable, rate-limited, or incompatible.
Try another permitted public instance or use `saved_html` mode with pages saved
from your browser.

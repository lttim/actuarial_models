"""Page-level renderers.

Each module in this package owns one logical page of the Streamlit app:

* ``overview.py`` -- product picker + landing
* ``pricing_run.py`` -- run a single pricing scenario
* ``what_if.py`` -- scenario sweeps and what-if analysis
* ``excel_replicator.py`` -- export-then-recompute parity gate
* ``alm.py`` -- ALM ladder run + projection
* ``unit_tests.py`` -- in-app pytest summary

Pages MUST import from :mod:`annuity_model` (the public surface), never
from internal modules. That way a future module rename does not ripple
into the UI layer.
"""

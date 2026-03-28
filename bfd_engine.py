"""
BFD Engine — Mass Balance Calculator
All quantities derived from molar ratios relative to SM batch size.
"""

class BFDEngine:

    def calculate(self, data):
        project    = data.get("project", {})
        components = data.get("components", [])
        operations = data.get("operations", [])

        batch_size = float(project.get("batch_size", 0))

        # Find SM (molar_ratio == 1 and role SM, or first component)
        sm = next((c for c in components if c.get("role","").upper() in ("SM","STARTING_MATERIAL")
                   and float(c.get("molar_ratio",1)) == 1), None)
        if not sm and components:
            sm = components[0]
        if not sm:
            return {"error": "No Starting Material defined"}

        sm_mw    = float(sm.get("mw", 1))
        sm_moles = batch_size / sm_mw          # kmol of SM

        # ── Enrich components ──────────────────────────────────────
        enriched = []
        for c in components:
            ec      = dict(c)
            mw      = float(c.get("mw", 1))
            mr      = float(c.get("molar_ratio", 1))
            purity  = float(c.get("purity", 100)) / 100
            density = float(c.get("density", 1000)) if c.get("density") else 1000

            moles       = sm_moles * mr
            mass_pure   = moles * mw
            mass_actual = mass_pure / purity if purity > 0 else mass_pure
            volume_l    = mass_actual / density          # litres
            volume_kl   = volume_l / 1000

            ec.update({
                "sm_moles"     : round(sm_moles, 4),
                "moles"        : round(moles, 4),
                "mass_kg"      : round(mass_actual, 3),
                "mass_pure_kg" : round(mass_pure, 3),
                "volume_kl"    : round(volume_kl, 4),
            })
            enriched.append(ec)

        result = {"components": enriched}

        # ── Stream counters ────────────────────────────────────────
        sc = {"IS": 0, "OS": 0, "PS": 0}
        def nid(t):
            sc[t] += 1
            return f"{t}-{sc[t]:02d}"

        # Running state passed between operations
        state = {}

        op_results = []
        for op in operations:
            op_type = op.get("type", "")
            res = {
                "type"              : op_type,
                "name"              : op.get("name", op_type.title()),
                "equipment"         : op.get("equipment", {}),
                "steps"             : op.get("steps", []),
                "temp_initial"      : op.get("temp_initial", ""),
                "temp_final"        : op.get("temp_final", ""),
                "pressure"          : op.get("pressure", "atm"),
                "inlet_streams"     : [],
                "outlet_streams"    : [],
                "process_stream"    : None,
                "operating_volume_kl": None,
            }

            # ── REACTION ──────────────────────────────────────────
            if op_type == "reaction":
                conv = float(op.get("conversion", 100)) / 100
                sel  = float(op.get("selectivity", 100)) / 100

                product = next((c for c in enriched if c.get("role","").lower() == "product"), None)
                product_mw   = float(product["mw"]) if product else sm_mw
                product_name = product["name"] if product else "Product"

                # Inlet streams: everything except product & byproducts
                for c in enriched:
                    role = c.get("role","").lower()
                    if role in ("product", "byproduct"):
                        continue
                    sid = nid("IS")
                    res["inlet_streams"].append({
                        "id": sid, "component": c["name"],
                        "qty_kg" : c["mass_kg"],
                        "vol_kl" : c["volume_kl"],
                    })

                # Reaction extents
                reacted_moles   = sm_moles * conv
                product_moles   = reacted_moles * sel
                product_mass    = product_moles * product_mw
                unreacted_sm    = (sm_moles - reacted_moles) * sm_mw
                solvent_kg      = sum(c["mass_kg"]  for c in enriched if c.get("role","").lower() == "solvent")
                solvent_vol     = sum(c["volume_kl"] for c in enriched if c.get("role","").lower() == "solvent")

                # Byproduct outlet streams
                for c in enriched:
                    if c.get("role","").lower() == "byproduct":
                        bp_moles = reacted_moles * float(c.get("molar_ratio", 0))
                        bp_mass  = bp_moles * float(c["mw"])
                        if bp_mass > 0.001:
                            sid = nid("OS")
                            density = float(c.get("density", 1.2))
                            res["outlet_streams"].append({
                                "id"       : sid,
                                "component": c["name"],
                                "qty_kg"   : round(bp_mass, 3),
                                "vol_kl"   : round(bp_mass / density / 1000, 4),
                                "category" : "gas_waste",
                            })

                # Process stream composition
                total_mass = product_mass + unreacted_sm + solvent_kg
                total_vol  = product_mass / product_mw * product_mw / 1000 + solvent_vol  # approx
                composition = []
                if product:
                    composition.append({"component": product_name,
                                        "qty_kg": round(product_mass, 3),
                                        "wt_pct": round(product_mass / total_mass * 100, 2) if total_mass else 0})
                for c in enriched:
                    if c.get("role","").lower() == "solvent":
                        composition.append({"component": c["name"],
                                            "qty_kg": c["mass_kg"],
                                            "wt_pct": round(c["mass_kg"] / total_mass * 100, 2) if total_mass else 0})
                if unreacted_sm > 0.01:
                    composition.append({"component": sm["name"] + " (unreacted)",
                                        "qty_kg": round(unreacted_sm, 3),
                                        "wt_pct": round(unreacted_sm / total_mass * 100, 2) if total_mass else 0})

                op_vol = total_vol + product_mass / (float(product.get("density",1000)) if product else 1000)
                res["operating_volume_kl"] = round(op_vol, 3)

                ps_id = nid("PS")
                res["process_stream"] = {
                    "id"         : ps_id,
                    "name"       : "Reaction Mass",
                    "qty_kg"     : round(total_mass, 3),
                    "vol_kl"     : round(op_vol, 4),
                    "composition": composition,
                }

                # Volume guardrail
                eq_vol = float(op.get("equipment", {}).get("volume_kl", 0))
                if eq_vol > 0 and op_vol > eq_vol * 0.8:
                    res["volume_warning"] = (f"⚠ Operating vol {op_vol:.2f} kL exceeds "
                                             f"80% of equipment vol {eq_vol} kL")

                state = {
                    "_product_name" : product_name,
                    "_product_mass" : product_mass,
                    "_product_moles": product_moles,
                    "_product_mw"   : product_mw,
                    "_unreacted_sm" : unreacted_sm,
                    "_solvent_kg"   : solvent_kg,
                    "_solvent_vol"  : solvent_vol,
                    "_total"        : total_mass,
                    "_sm_name"      : sm["name"],
                    "_sm_mw"        : sm_mw,
                    "_sm_moles"     : sm_moles,
                }

            # ── DISTILLATION ──────────────────────────────────────
            elif op_type == "distillation":
                dist_frac   = float(op.get("distillate_fraction", 0.95))
                sol_kg      = state.get("_solvent_kg", 0)
                sol_distill = sol_kg * dist_frac
                sol_bottom  = sol_kg * (1 - dist_frac)

                prev_ps = op_results[-1]["process_stream"] if op_results else None
                if prev_ps:
                    res["inlet_streams"].append({
                        "id": prev_ps["id"], "component": prev_ps["name"],
                        "qty_kg": prev_ps["qty_kg"], "vol_kl": prev_ps["vol_kl"],
                    })

                sid = nid("OS")
                res["outlet_streams"].append({
                    "id": sid, "component": "Solvent Distillate (Recovered)",
                    "qty_kg"  : round(sol_distill, 3),
                    "vol_kl"  : round(sol_distill / 1000, 4),
                    "category": "organic_waste",
                    "note"    : "Recyclable",
                })

                product_mass  = state.get("_product_mass", 0)
                product_name  = state.get("_product_name", "Product")
                bottoms_mass  = product_mass + state.get("_unreacted_sm", 0) + sol_bottom
                ps_id = nid("PS")
                res["process_stream"] = {
                    "id"  : ps_id, "name": "Distillation Bottoms",
                    "qty_kg": round(bottoms_mass, 3),
                    "vol_kl": round(bottoms_mass / 1000, 4),
                    "composition": [
                        {"component": product_name,      "qty_kg": round(product_mass, 3),
                         "wt_pct": round(product_mass / bottoms_mass * 100, 2) if bottoms_mass else 0},
                        {"component": "Residual Solvent", "qty_kg": round(sol_bottom, 3),
                         "wt_pct": round(sol_bottom / bottoms_mass * 100, 2) if bottoms_mass else 0},
                    ],
                }
                state["_solvent_kg"] = sol_bottom
                state["_total"]      = bottoms_mass

            # ── FILTRATION ────────────────────────────────────────
            elif op_type == "filtration":
                lod          = float(op.get("lod", 30)) / 100
                wash_ratio   = float(op.get("wash_ratio", 2))
                product_loss = float(op.get("product_loss", 2)) / 100

                product_mass = state.get("_product_mass", 0)
                product_name = state.get("_product_name", "Product")
                solvent_kg   = state.get("_solvent_kg", 0)

                product_in_cake = product_mass * (1 - product_loss)
                product_in_ml   = product_mass * product_loss
                wet_cake_mass   = product_in_cake / (1 - lod) if lod < 1 else product_in_cake
                moisture_in_cake = wet_cake_mass - product_in_cake
                wash_solvent_kg  = product_in_cake * wash_ratio

                prev_ps = op_results[-1]["process_stream"] if op_results else None
                if prev_ps:
                    res["inlet_streams"].append({
                        "id": prev_ps["id"], "component": prev_ps["name"],
                        "qty_kg": prev_ps["qty_kg"], "vol_kl": prev_ps["vol_kl"],
                    })
                ws_id = nid("IS")
                res["inlet_streams"].append({
                    "id": ws_id, "component": "Wash Solvent",
                    "qty_kg": round(wash_solvent_kg, 3),
                    "vol_kl": round(wash_solvent_kg / 1000, 4),
                })

                # Mother liquor OS
                ml_mass = solvent_kg + product_in_ml
                ml_id   = nid("OS")
                res["outlet_streams"].append({
                    "id": ml_id, "component": "Mother Liquor",
                    "qty_kg": round(ml_mass, 3), "vol_kl": round(ml_mass / 1000, 4),
                    "category": "organic_waste",
                    "composition": [
                        {"component": "Solvent",                        "qty_kg": round(solvent_kg, 3)},
                        {"component": product_name + " (dissolved)",    "qty_kg": round(product_in_ml, 3)},
                    ],
                })
                # Wash liquor OS
                wl_id = nid("OS")
                res["outlet_streams"].append({
                    "id": wl_id, "component": "Wash Liquor",
                    "qty_kg": round(wash_solvent_kg * 1.05, 3),
                    "vol_kl": round(wash_solvent_kg * 1.05 / 1000, 4),
                    "category": "aqueous_waste",
                })

                # Wet cake PS
                ps_id = nid("PS")
                res["process_stream"] = {
                    "id": ps_id, "name": "Wet Cake",
                    "qty_kg": round(wet_cake_mass, 3),
                    "vol_kl": round(wet_cake_mass / 500, 4),
                    "composition": [
                        {"component": product_name, "qty_kg": round(product_in_cake, 3),
                         "wt_pct": round((1 - lod) * 100, 2)},
                        {"component": "Moisture",   "qty_kg": round(moisture_in_cake, 3),
                         "wt_pct": round(lod * 100, 2)},
                    ],
                }
                state["_product_mass"] = product_in_cake
                state["_wet_cake"]     = wet_cake_mass
                state["_total"]        = wet_cake_mass

            # ── DRYING ────────────────────────────────────────────
            elif op_type == "drying":
                lod_i = float(op.get("lod_initial", 30)) / 100
                lod_f = float(op.get("lod_final",   0.5)) / 100

                product_in_cake = state.get("_product_mass", 0)
                wet_cake        = state.get("_wet_cake", product_in_cake / (1 - lod_i) if lod_i < 1 else product_in_cake)
                product_name    = state.get("_product_name", "Product")

                dry_mass    = product_in_cake / (1 - lod_f) if lod_f < 1 else product_in_cake
                evaporated  = wet_cake - dry_mass

                prev_ps = op_results[-1]["process_stream"] if op_results else None
                if prev_ps:
                    res["inlet_streams"].append({
                        "id": prev_ps["id"], "component": prev_ps["name"],
                        "qty_kg": prev_ps["qty_kg"], "vol_kl": prev_ps["vol_kl"],
                    })

                cond_id = nid("OS")
                res["outlet_streams"].append({
                    "id": cond_id, "component": "Condensate",
                    "qty_kg": round(max(evaporated, 0), 3),
                    "vol_kl": round(max(evaporated, 0) / 1000, 4),
                    "category": "aqueous_waste",
                })

                ps_id = nid("PS")
                res["process_stream"] = {
                    "id": ps_id, "name": "Dry Product",
                    "qty_kg": round(dry_mass, 3),
                    "vol_kl": round(dry_mass / 500, 4),
                    "composition": [
                        {"component": product_name,        "qty_kg": round(product_in_cake, 3),
                         "wt_pct": round((1 - lod_f) * 100, 2)},
                        {"component": "Residual Moisture", "qty_kg": round(dry_mass - product_in_cake, 3),
                         "wt_pct": round(lod_f * 100, 2)},
                    ],
                }
                state["_dry_product"]   = dry_mass
                state["_final_product"] = product_in_cake

            op_results.append(res)

        result["operations"] = op_results

        # ── Overall yield ──────────────────────────────────────────
        final_product = state.get("_final_product", state.get("_product_mass", 0))
        product_mw    = state.get("_product_mw", sm_mw)
        if final_product > 0 and product_mw > 0:
            molar_yield = (final_product / product_mw) / sm_moles * 100 if sm_moles else 0
        else:
            molar_yield = 0

        result["yield"] = {
            "sm_input_kg"       : round(batch_size, 3),
            "product_output_kg" : round(final_product, 3),
            "molar_yield_pct"   : round(molar_yield, 2),
        }

        return result

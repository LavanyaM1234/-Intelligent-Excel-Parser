"""Parameter and Asset registries for the Excel parser"""

# Known parameters with metadata
PARAMETER_REGISTRY = [
    {
        "name": "coal_consumption",
        "display_name": "Coal Consumption",
        "unit": "MT",
        "category": "input",
        "section": "COGEN BOILER",
        "aliases": ["Coal Consumption", "Coal Used", "Coal (MT)", "Daily Coal", "Coal Used (MT)", "COAL CONSMPTN"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "coal_gcv",
        "display_name": "Coal GCV",
        "unit": "kcal/kg",
        "category": "input",
        "section": "COGEN BOILER",
        "aliases": ["Coal GCV", "Coal Gross Calorific Value", "GCV", "COALCV"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "steam_generation",
        "display_name": "Steam Generation",
        "unit": "T/hr",
        "category": "output",
        "section": "COGEN BOILER",
        "aliases": ["Steam", "Steam Generated", "Steam Generation", "Steam (Boiler)", "Steam Output", "STEAM GEN"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "steam_consumption",
        "display_name": "Steam Consumption",
        "unit": "T/hr",
        "category": "input",
        "section": "COGEN BOILER",
        "aliases": ["Steam Consumption", "Steam Used", "Steam Input"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "power_generation",
        "display_name": "Power Generation",
        "unit": "MWh",
        "category": "output",
        "section": "POWER PLANT",
        "aliases": ["Power", "Power Generated", "Power Output", "Power (MW)", "Power Generation", "Power TG", "POWER GEN"],
        "applicable_assets": ["TG-1", "TG-2"]
    },
    {
        "name": "power_consumption",
        "display_name": "Power Consumption",
        "unit": "MWh",
        "category": "input",
        "section": "POWER PLANT",
        "aliases": ["Power Consumption", "Power Used", "Auxiliary Power Load"],
        "applicable_assets": ["TG-1", "TG-2"]
    },
    {
        "name": "power_export",
        "display_name": "Power Export",
        "unit": "MWh",
        "category": "output",
        "section": "POWER PLANT",
        "aliases": ["Power Export", "Power Exported", "Grid Export"],
        "applicable_assets": ["TG-1", "TG-2"]
    },
    {
        "name": "water_consumption",
        "display_name": "Water Consumption",
        "unit": "KL",
        "category": "input",
        "section": "UTILITIES",
        "aliases": ["Water", "Water Consumption", "Water Used", "Water (KL)"],
        "applicable_assets": ["AFBC-1", "AFBC-2", "TG-1", "TG-2"]
    },
    {
        "name": "co2_emissions",
        "display_name": "CO₂ Emissions",
        "unit": "tCO2e",
        "category": "emission",
        "section": "EMISSIONS",
        "aliases": ["CO2 Emissions", "CO2", "Carbon Dioxide"],
        "applicable_assets": ["AFBC-1", "AFBC-2", "TG-1", "TG-2"]
    },
    {
        "name": "so2_emissions",
        "display_name": "SO₂ Emissions",
        "unit": "kg",
        "category": "emission",
        "section": "EMISSIONS",
        "aliases": ["SO2 Emissions", "SO2", "Sulfur Dioxide"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "nox_emissions",
        "display_name": "NOx Emissions",
        "unit": "kg",
        "category": "emission",
        "section": "EMISSIONS",
        "aliases": ["NOx Emissions", "NOx", "Nitrogen Oxides"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "fly_ash_generated",
        "display_name": "Fly Ash Generated",
        "unit": "MT",
        "category": "output",
        "section": "WASTE",
        "aliases": ["Fly Ash", "Fly Ash Generated", "Ash Output"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "efficiency",
        "display_name": "Boiler Efficiency",
        "unit": "%",
        "category": "calculated",
        "section": "COGEN BOILER",
        "aliases": ["Efficiency", "Plant Efficiency", "Overall Efficiency", "EFF %", "Boiler Efficiency"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "specific_coal_consumption",
        "display_name": "Specific Coal Consumption",
        "unit": "kg/kWh",
        "category": "calculated",
        "section": "COGEN BOILER",
        "aliases": ["Specific Coal Consumption", "SCC", "Coal Per Unit"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "heat_rate",
        "display_name": "Heat Rate",
        "unit": "kcal/kWh",
        "category": "calculated",
        "section": "POWER PLANT",
        "aliases": ["Heat Rate", "Heat Input Rate", "HR"],
        "applicable_assets": ["TG-1", "TG-2"]
    },
    {
        "name": "plant_load_factor",
        "display_name": "Plant Load Factor",
        "unit": "%",
        "category": "calculated",
        "section": "POWER PLANT",
        "aliases": ["Plant Load Factor", "PLF", "Capacity Factor"],
        "applicable_assets": ["TG-1", "TG-2"]
    },
    {
        "name": "lignite_consumption",
        "display_name": "Lignite Consumption",
        "unit": "MT",
        "category": "input",
        "section": "COGEN BOILER",
        "aliases": ["Lignite", "Lignite Consumption", "Brown Coal"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "biomass_consumption",
        "display_name": "Biomass Consumption",
        "unit": "MT",
        "category": "input",
        "section": "COGEN BOILER",
        "aliases": ["Biomass", "Biomass Consumption", "Bio-fuel"],
        "applicable_assets": ["AFBC-1", "AFBC-2"]
    },
    {
        "name": "production_output",
        "display_name": "Production Output",
        "unit": "MT",
        "category": "output",
        "section": "PRODUCTION",
        "aliases": ["Production", "Output", "Production Output"],
        "applicable_assets": ["VSF"]
    },
    {
        "name": "operating_hours",
        "display_name": "Operating Hours",
        "unit": "hrs",
        "category": "input",
        "section": "OPERATIONS",
        "aliases": ["Operating Hours", "Runtime", "Hours", "HOURS", "Uptime"],
        "applicable_assets": ["AFBC-1", "AFBC-2", "TG-1", "TG-2", "VSF"]
    }
]

# Known assets
ASSET_REGISTRY = [
    {
        "name": "AFBC-1",
        "display_name": "AFBC Boiler 1",
        "type": "boiler",
        "aliases": ["AFBC-1", "Boiler 1", "AFBC 1", "AFB1"]
    },
    {
        "name": "AFBC-2",
        "display_name": "AFBC Boiler 2",
        "type": "boiler",
        "aliases": ["AFBC-2", "Boiler 2", "AFBC 2", "AFB2"]
    },
    {
        "name": "TG-1",
        "display_name": "Turbo Generator 1",
        "type": "turbine",
        "aliases": ["TG-1", "TG1", "Turbine 1", "Generator 1"]
    },
    {
        "name": "TG-2",
        "display_name": "Turbo Generator 2",
        "type": "turbine",
        "aliases": ["TG-2", "TG2", "Turbine 2", "Generator 2"]
    },
    {
        "name": "VSF",
        "display_name": "Viscose Staple Fiber",
        "type": "product",
        "aliases": ["VSF", "Viscose Staple Fiber", "Fiber"]
    },
    {
        "name": "KILN-1",
        "display_name": "Rotary Kiln 1",
        "type": "kiln",
        "aliases": ["KILN-1", "Kiln 1", "Rotary Kiln 1", "RK1"]
    }
]


# Helper functions to work with registries

def get_parameter_by_name(param_name: str) -> dict:
    """Get parameter details by name"""
    for param in PARAMETER_REGISTRY:
        if param["name"] == param_name:
            return param
    return None


def get_asset_by_name(asset_name: str) -> dict:
    """Get asset details by name"""
    for asset in ASSET_REGISTRY:
        if asset["name"] == asset_name:
            return asset
    return None


def get_parameters_by_section(section: str) -> list:
    """Get all parameters in a specific section"""
    return [p for p in PARAMETER_REGISTRY if p["section"] == section]


def get_parameters_by_category(category: str) -> list:
    """Get all parameters by category (input, output, calculated, emission)"""
    return [p for p in PARAMETER_REGISTRY if p["category"] == category]


def get_parameters_for_asset(asset_name: str) -> list:
    """Get all parameters applicable to a specific asset"""
    return [p for p in PARAMETER_REGISTRY if asset_name in p.get("applicable_assets", [])]


def get_assets_by_type(asset_type: str) -> list:
    """Get all assets of a specific type"""
    return [a for a in ASSET_REGISTRY if a["type"] == asset_type]


def get_all_parameter_names() -> list:
    """Get list of all parameter names"""
    return [p["name"] for p in PARAMETER_REGISTRY]


def get_all_asset_names() -> list:
    """Get list of all asset names"""
    return [a["name"] for a in ASSET_REGISTRY]


def get_parameter_context() -> str:
    """Generate context string for LLM with parameter registry"""
    context = "# Parameter Registry\n"
    for param in PARAMETER_REGISTRY:
        aliases = ", ".join(param["aliases"])
        context += f"\n- {param['name']}: {param['display_name']} ({param['unit']})\n"
        context += f"  Category: {param['category']}\n"
        context += f"  Section: {param['section']}\n"
        context += f"  Aliases: {aliases}\n"
        if param.get("applicable_assets"):
            assets = ", ".join(param["applicable_assets"])
            context += f"  Applicable Assets: {assets}\n"
    return context


def get_asset_context() -> str:
    """Generate context string for LLM with asset registry"""
    context = "# Asset Registry\n"
    for asset in ASSET_REGISTRY:
        aliases = ", ".join(asset["aliases"])
        context += f"\n- {asset['name']}: {asset['display_name']} ({asset['type']})\n"
        context += f"  Aliases: {aliases}\n"
    return context


def validate_parameter_for_asset(param_name: str, asset_name: str) -> bool:
    """Check if parameter is applicable to asset"""
    param = get_parameter_by_name(param_name)
    if not param:
        return False
    return asset_name in param.get("applicable_assets", [])


def get_parameter_unit(param_name: str) -> str:
    """Get unit for a parameter"""
    param = get_parameter_by_name(param_name)
    return param["unit"] if param else None


def get_all_sections() -> list:
    """Get all unique sections"""
    sections = set()
    for param in PARAMETER_REGISTRY:
        sections.add(param["section"])
    return sorted(list(sections))


def export_registries_as_json() -> dict:
    """Export registries in JSON format"""
    return {
        "parameters": PARAMETER_REGISTRY,
        "assets": ASSET_REGISTRY
    }


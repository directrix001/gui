"""
Aluminium part price calculator.

Formula:
    AMS_Q      = (MC_Q   × DF_c) + CNG_Q
    AMS_Q-1    = (MC_Q-1 × DF_c) + CNG_Q-1
    PPI_Factor = (PPI_Q − PPI_Q-1) / PPI_Q-1
    P_New      = P_Current + [(AMS_Q − AMS_Q-1) × PWt] + (PPI_Factor × P_Current)
"""
from dataclasses import dataclass


@dataclass
class PriceInputs:
    weight: float           # PWt
    current_price: float    # P_Current
    ppi_q: float
    ppi_q1: float           # PPI_Q-1
    drauss_factor: float    # DF_c (default 1.44 upstream)
    mc_q: float
    mc_q_1: float
    cng_q: float
    cng_q_1: float


@dataclass
class PriceResult:
    new_price: float
    ams_q: float
    ams_q_1: float
    ppi_factor: float


def calculate_new_price(inputs: PriceInputs) -> PriceResult:
    if inputs.ppi_q1 == 0:
        raise ValueError("PPI_Q-1 cannot be zero (division by zero in PPI factor).")

    ams_q = (inputs.mc_q * inputs.drauss_factor) + inputs.cng_q
    ams_q_1 = (inputs.mc_q_1 * inputs.drauss_factor) + inputs.cng_q_1
    ppi_factor = (inputs.ppi_q - inputs.ppi_q1) / inputs.ppi_q1
    new_price = (
        inputs.current_price
        + ((ams_q - ams_q_1) * inputs.weight)
        + (ppi_factor * inputs.current_price)
    )
    return PriceResult(
        new_price=round(new_price, 4),
        ams_q=round(ams_q, 4),
        ams_q_1=round(ams_q_1, 4),
        ppi_factor=round(ppi_factor, 6),
    )

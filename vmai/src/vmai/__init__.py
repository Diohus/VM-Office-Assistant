"""API công khai của vmai. Xem docs/contracts/vmai_api.md.

Import gói này không kéo theo SDK LLM và không đọc biến môi trường key.
"""
from .confidence import RULE_VERSION, Color, grade_formula
from .contract import SCHEMA_VERSION, LlmCall, ProposalBatch
from .confirm import ConfirmRow, SourceOption, confirmation_rows, source_options
from .errors import FeatureNotReleased, MissingFields, SourceFileError, TemplateSpecError, VmaiError
from .headers import HeaderPlan, plan_headers
from .infer import SpecGuess, infer_spec, profile_from_guess
from .llm_tasks import map_headers
from .phi import ExternalSendPolicy, PhiDetected, SendNotAllowed, Token, guard_external_image, scan_tokens
from .profile import ProfileStatus, TemplateProfile
from .propose import propose
from .security import FormulaCheck, FormulaRisk, check_formula
from .source_xlsx import SourceFacts, read_source
from .template import FormFacts, TemplateSpec, load_spec, read_form

__all__ = [
    "RULE_VERSION",
    "SCHEMA_VERSION",
    "Color",
    "ExternalSendPolicy",
    "ConfirmRow",
    "FeatureNotReleased",
    "SourceOption",
    "confirmation_rows",
    "source_options",
    "FormFacts",
    "FormulaCheck",
    "FormulaRisk",
    "HeaderPlan",
    "LlmCall",
    "MissingFields",
    "PhiDetected",
    "ProfileStatus",
    "ProposalBatch",
    "SendNotAllowed",
    "SourceFacts",
    "SpecGuess",
    "SourceFileError",
    "TemplateProfile",
    "TemplateSpec",
    "TemplateSpecError",
    "Token",
    "VmaiError",
    "check_formula",
    "grade_formula",
    "guard_external_image",
    "infer_spec",
    "load_spec",
    "map_headers",
    "plan_headers",
    "profile_from_guess",
    "propose",
    "read_form",
    "read_source",
    "scan_tokens",
]

# Convenience re-exports for module functions
from .template_loader import TemplateRepository, TemplateNotFoundError
from .data_prep import DataPrepResult, prepare_facts_for_template
from .llm_orchestrator import LLMOrchestrator
from .validator import ValidationResult, Validator
from .ppt_generator import PPTGenerator
from .audit_logger import AuditLogger

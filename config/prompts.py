from dataclasses import dataclass

@dataclass
class PromptHolder:
    
    STRUCTURE_SCHEMA_PROMPT : str = """
    Return a strict JSON object with these fields ONLY (no additional keys anywhere):
    - name: string
    - contact: { email: string|null, phone: string|null, location: string|null, links: string[] }
    - summary: string|null
    - experience: [ { title: string, company: string|null, location: string|null, start_date: string|null, end_date: string|null, achievements: string[] } ]
    - education: [ { degree: string|null, institution: string|null, location: string|null, start_date: string|null, end_date: string|null, gpa: string|null } ]
    - skills: { technical: string[], tools: string[], soft: string[] }
    - certifications: string[]
    - projects: [ { name: string, description: string, technologies: string[] } ]
    - awards: string[]
    - languages: string[]
    Constraints:
    - Output ONLY a single JSON object. No prose, no markdown, no code fences.
    - If a value is unknown, use null (for scalars) or [] (for arrays). Do NOT fabricate.
    - Use only the keys shown above; do NOT add other keys or nested structures.
    - contact.links must be fully-qualified URLs when present. Strip trailing punctuation.
    - Preserve original wording where possible; do minor normalization only.
    - Dates may be free-form (e.g., "Jan 2021", "Present").
    """

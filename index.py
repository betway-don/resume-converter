import pdfplumber
import json
import os
from openai import OpenAI
from docxtpl import DocxTemplate

# --- CONFIGURATION ---
OPENAI_API_KEY = "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"
TEMPLATE_PATH = "doom.docx"
OUTPUT_PATH = "Final_Generated_Resume.docx"
PDF_PATH = "candidate_resume.pdf"

client = OpenAI(api_key=OPENAI_API_KEY)

def extract_text_from_pdf(pdf_path):
    """Reads raw text from the candidate's PDF"""
    print(f"Reading PDF: {pdf_path}...")
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text

def get_ai_data(resume_text):
    """Sends text to OpenAI and asks for structured JSON"""
    print("Sending data to AI for extraction...")
    
    system_prompt = """
    You are a professional resume parser following Zensar resume guidelines.
    
    CRITICAL RULES:
    1. Summary MUST be 5+ lines highlighting total experience, projects, expertise
    2. Skills MUST be an ARRAY where category names are DYNAMIC based on candidate's profile
       - Examples: "Domain Experience", "Gaming Platforms", "Cloud Technologies", "Tools & Frameworks"
       - Each skill object has: category (string), primary (string), secondary (string)
    3. Experience summary MUST have 2-3 achievement bullets per role
    4. Quantify achievements with numbers, percentages, or specific examples
    5. Avoid generic phrases like "collaborated with team", "resolved bugs"
    6. Extract ONLY relevant certifications and awards (no school-level)
    7. Education should have passing year after qualification
    8. Limit to 3-5 key projects most relevant to the role
    """

    user_prompt = """
    Resume Content:
    """ + resume_text + """

    Extract and structure the data in this EXACT JSON format:
    {
      "personal_info": {
        "name": "Full Name",
        "email": "email@example.com",
        "phone": "+1234567890",
        "location": "City, Country"
      },
      "summary": "5+ line summary highlighting: total years of experience, number of projects/accounts worked on, diverse topics or geographies covered, core expertise areas. Should be crisp and impactful.",
      "education": [
        {"degree": "Bachelor of Engineering in Computer Science", "year": "2018"}
      ],
      "skills": [
        {
          "category": "Domain Experience",
          "primary": "Banking, Insurance, Financial Services",
          "secondary": "Healthcare, Retail"
        },
        {
          "category": "Tools & Frameworks",
          "primary": "JIRA, Git, Jenkins, Angular",
          "secondary": "Docker, Kubernetes, React"
        },
        {
          "category": "Databases",
          "primary": "MySQL, Oracle, SQL Server",
          "secondary": "MongoDB, PostgreSQL"
        },
        {
          "category": "Defect Tracking",
          "primary": "BugZilla, Mantis",
          "secondary": "Redmine, Jira Service Desk"
        },
        {
          "category": "Programming Languages",
          "primary": "JavaScript, TypeScript, Python",
          "secondary": "Java, C#"
        }
      ],
      "expertise_areas": ["Expertise Area 1", "Expertise Area 2", "Expertise Area 3"],
      "certifications": ["AWS Certified Solutions Architect", "Certified Scrum Master"],
      "awards": ["Employee of the Year 2023", "Innovation Award 2022"],
      "experience_summary": [
        {
          "role": "Senior Software Engineer",
          "years": "2022-Present",
          "skills": "Angular, WebView, Microservices",
          "achievements": [
            "Developed micro frontend application reducing load time by 40%",
            "Migrated 5 projects to Angular v18-19, improving performance by 30%",
            "Mentored 3 junior developers, reducing code review cycles by 25%"
          ]
        },
        {
          "role": "Web UI Developer",
          "years": "2020-2021",
          "skills": "JavaScript, Angular 8, PrimeNG",
          "achievements": [
            "Designed and delivered 12 new UI widgets for enterprise dashboard",
            "Ensured 100% WCAG 2.1 AA compliance across 15+ components"
          ]
        }
      ]
    }

    IMPORTANT EXTRACTION RULES:
    - For skills: Create categories DYNAMICALLY based on candidate's actual experience
      * If candidate works in gaming: use "Gaming Platforms", "Game Engines"
      * If candidate works in cloud: use "Cloud Platforms", "DevOps Tools"
      * If candidate works in web dev: use "Frontend Frameworks", "Backend Technologies"
    - Split each category into PRIMARY (most used/expert) and SECONDARY (familiar/supporting)
    - For achievements: Include specific numbers, percentages, or quantifiable impact
    - For experience_summary: Limit to 3-5 most relevant projects
    - For education: Include only engineering degree and above with passing year
    - For certifications/awards: Only include professional/relevant ones
    - Summary must be minimum 5 lines covering experience breadth and depth
    """

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        response_format={"type": "json_object"}
    )
    
    return json.loads(response.choices[0].message.content)

def generate_doc(data, template_file, output_file):
    """Injects the JSON data into the Word Template"""
    print("Generating Word Document...")
    
    # The data is already in the correct format for the template
    # Skills is now an array of objects with category, primary, secondary
    
    doc = DocxTemplate(template_file)
    doc.render(data)
    doc.save(output_file)
    print(f"Success! Resume saved to: {output_file}")

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    # 1. Check if files exist
    if not os.path.exists(TEMPLATE_PATH):
        print(f"Error: Template file '{TEMPLATE_PATH}' not found.")
    elif not os.path.exists(PDF_PATH):
        print(f"Error: PDF file '{PDF_PATH}' not found.")
    else:
        # 2. Extract Text
        raw_text = extract_text_from_pdf(PDF_PATH)
        
        # 3. Get Structured Data
        resume_data = get_ai_data(raw_text)
        
        # 4. Print data to verify extraction
        print("\n" + "="*80)
        print("EXTRACTED DATA:")
        print("="*80)
        print(json.dumps(resume_data, indent=2))
        print("="*80 + "\n")
        
        # 5. Generate Word Doc
        generate_doc(resume_data, TEMPLATE_PATH, OUTPUT_PATH)
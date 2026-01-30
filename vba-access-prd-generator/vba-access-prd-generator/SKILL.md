---
name: vba-access-prd-generator
description: Generates comprehensive Product Requirements Documents (PRDs) for VBA Microsoft Access applications by analyzing exported source code, database schemas, and UI. Use when the user wants to document an existing Access application for migration to modern web technologies, needs to create PRDs from legacy code, or wants to understand form functionality, data flows, navigation patterns, and business logic in Access databases. Detects unused forms and dead code.
---

# VBA Access PRD Generator

This skill generates exhaustive Product Requirements Documents (PRDs) from VBA Microsoft Access applications by performing reverse engineering of the codebase, database structure, and user interface.

## Workflow Overview

The PRD generation follows an interactive, iterative process:

1. **Initial Analysis** - Analyze all source files to build a complete inventory
2. **Form-by-Form Deep Dive** - Interactive review of each form with user input
3. **Data Model Documentation** - Process ERD and database relationships
4. **Navigation & Dependencies** - Map form flows and module dependencies
5. **Code Quality Assessment** - Identify dead code and refactoring opportunities
6. **PRD Generation** - Create master PRD and individual module PRDs
7. **User Stories & API Specs** - Infer requirements and technical specifications

## Prerequisites

The user should provide:
- **Source code folder** (`src/` directory) containing exported VBA files:
  - Form modules: `Form_*.cls` files
  - Standard modules: `*.bas` files
  - Class modules: `*.cls` files
- **ERD file** - Database schema in any format (text, image, SQL DDL)
- **Willingness to interact** - User will be asked for screenshots and clarifications

## Step 1: Initial Project Setup

Start by understanding the project context:

```markdown
I'll help you generate comprehensive PRDs for your Access application. Let me start by asking a few questions:

1. What is the name of this application?
2. What is its main business purpose?
3. Who are the primary users (roles)?
4. Where is the src/ folder located with the exported VBA files?
5. Do you have an ERD or database schema file?
```

Once confirmed, proceed to analyze all files.

## Step 2: Comprehensive Code Analysis

Use the analysis script to scan all VBA files:

```bash
python3 scripts/analyze_form.py /path/to/src > analysis_results.json
```

This generates a complete inventory of:
- All forms with their controls and functionality
- All modules and their functions
- Form navigation dependencies
- Database tables accessed
- Potential dead code
- Unused forms

Present a summary to the user:

```markdown
## Analysis Complete

I've analyzed your application and found:
- **X forms** (Y appear to be unused)
- **Z standard modules** (including N utility modules)
- **W class modules**
- **Navigation flow** detected between forms
- **Potential dead code** in M locations

I'll now walk through each form interactively. For forms with complex UI, I'll ask you to provide screenshots.
```

## Step 3: Interactive Form Documentation

For each form, follow this process:

### 3.1 Present Initial Analysis

```markdown
## Form: [FormName]

**Detected Controls:**
- [List all controls found]

**Event Handlers:**
- [List all event handlers]

**Data Sources:**
- Tables: [List tables accessed]
- SQL Queries: [Show main queries]

**Navigation:**
- Opens: [Forms called from this form]
- Opened by: [Forms that call this form]

**Questions:**
1. Can you provide a screenshot of this form to understand its layout?
2. What is the primary purpose of this form from a user perspective?
3. Are there any hidden business rules or validations not obvious in the code?
4. Is this form still actively used?
```

### 3.2 Gather Screenshot and Context

If the user provides a screenshot:
- Analyze the visual layout
- Map controls to their visible labels
- Understand the user flow
- Identify groupings and tabs

### 3.3 Infer User Stories

Based on code analysis and user input, generate user stories:

```markdown
### Inferred User Stories for [FormName]:

**Story 1: [Action Name]**
As a [user role]
I want to [action based on form functionality]
So that [inferred business benefit]

Acceptance Criteria:
- [ ] [Derived from validation code]
- [ ] [Derived from business logic]
- [ ] [Derived from data operations]

**Technical Notes:**
- VBA Implementation: [How it's currently done]
- Data Operations: [Tables/queries involved]
- Modern Equivalent: [Recommended approach]
```

### 3.4 Document API Requirements

For each CRUD operation detected:

```markdown
### Required API Endpoints for [FormName]:

#### GET /api/[resource]/{id}
**Purpose:** Retrieve single [entity] record
**VBA Origin:** RecordSource or OpenRecordset in Form_Load
**Request:** 
- Path param: id (integer)
**Response:**
```json
{
  "field1": "type",
  "field2": "type"
}
```

#### POST /api/[resource]
**Purpose:** Create new [entity]
**VBA Origin:** AddNew/Update in Save button click
**Request Body:**
[Based on form fields]
**Business Rules:**
- [From validation code]
```

## Step 4: Data Model Documentation

Process the ERD file provided:

```markdown
## Data Model Analysis

I've analyzed the database schema from your ERD. Let me document:

1. **Core Entities:**
   - [List main tables with their purpose]

2. **Relationships:**
   - [Key relationships with cardinality]

3. **Data Volumes:**
   - Are there any tables with significant data volume concerns?
   - Any performance issues with current queries?

4. **Migration Considerations:**
   - [Normalization opportunities]
   - [Index recommendations]
   - [Data type optimizations]
```

Use the ERD to enhance form documentation by linking forms to their data entities.

## Step 5: Generate Navigation Flow

Create visual navigation diagram:

```bash
python3 scripts/generate_flow.py analysis_results.json > navigation_flow.md
```

Include in master PRD:

```markdown
## Application Navigation Flow

The following diagram shows how users navigate between forms:

```mermaid
[Generated diagram]
```

**Entry Points:**
- [Forms that are never called - likely menu/startup forms]

**Most Connected Forms:**
- [Forms with most incoming/outgoing connections]

**Isolated Forms:**
- [Forms with no connections - candidates for removal]
```

## Step 6: Code Quality & Dead Code Analysis

Document findings:

```markdown
## Code Quality Assessment

### Unused Forms (Candidates for Removal)
The following forms are never opened by any other form:
- [FormName1] - Last modified: [if available]
- [FormName2]

**Recommendation:** Verify these are not entry points before removing.

### Dead Code Suspects
| Location | Type | Name | Reason |
|----------|------|------|--------|
| [Form] | Function | [Name] | [Reason] |

### Refactoring Opportunities
1. **Duplicate Code:**
   - [Functions appearing in multiple forms]
   - Recommendation: Extract to shared module

2. **Complex Functions:**
   - [Functions over 100 lines]
   - Recommendation: Break down into smaller units

3. **Hardcoded Values:**
   - [Magic numbers/strings in code]
   - Recommendation: Move to configuration

4. **SQL Injection Risks:**
   - [String concatenation in SQL]
   - Recommendation: Use parameterized queries in new system
```

## Step 7: Generate PRD Documents

Create two levels of documentation:

### 7.1 Master PRD

Use template from `references/prd_templates.md` to create comprehensive master document:

```markdown
# PRD: [Application Name]

[Following the Main PRD Structure template]

Include:
- Executive summary
- System overview
- Links to all module PRDs
- Complete data model
- Navigation flows
- Migration strategy
- Technical debt assessment
```

### 7.2 Individual Module PRDs

For each significant form/module:

```markdown
# PRD Module: [FormName]

[Following the Module PRD Structure template]

Include:
- Screenshot (if provided)
- Control inventory
- User stories
- API specifications
- Business rules
- Migration considerations
```

## Step 8: Technology Stack Recommendations

Based on the analysis, recommend appropriate modern stack:

```markdown
## Recommended Technology Stack

### Backend
- **Framework:** [Recommendation based on complexity]
- **Database:** [Based on data model]
- **Key Reasons:** [Why this stack fits]

### Frontend
- **Framework:** [React/Vue/Angular]
- **UI Library:** [Based on form complexity]
- **Key Reasons:** [Why this stack fits]

### Migration Phases
**Phase 1:** [Critical path forms]
**Phase 2:** [Secondary features]
**Phase 3:** [Nice-to-have features]
```

Consult `references/migration_patterns.md` for best practices.

## Output Structure

Generate the following files in a `prd-output/` folder:

```
prd-output/
├── PRD_Master_[AppName].md                    # Main PRD
├── modules/
│   ├── PRD_Module_[FormName1].md              # Individual form PRDs
│   ├── PRD_Module_[FormName2].md
│   └── ...
├── diagrams/
│   ├── navigation_flow.md                     # Mermaid diagrams
│   ├── data_model.md
│   └── module_dependencies.md
├── technical/
│   ├── api_specifications.md                  # All API endpoints
│   ├── user_stories.md                        # All user stories
│   └── code_quality_report.md                 # Dead code & refactoring
└── analysis_results.json                      # Raw analysis data
```

## Interactive Prompts

Throughout the process, ask clarifying questions:

**For each form:**
- "Can you provide a screenshot of [FormName]?"
- "What role uses this form primarily?"
- "Is [detected functionality] still relevant?"

**For unclear code:**
- "I see [code pattern]. What is the business purpose?"
- "This validation seems complex. Can you explain the rule?"

**For dead code:**
- "Form [X] is never called. Is it still needed?"
- "Function [Y] is never used. Safe to document as removable?"

## Best Practices

1. **Be Thorough:** Don't skip forms even if they seem simple
2. **Ask Questions:** Better to over-clarify than make assumptions
3. **Visual Context:** Always request screenshots for complex forms
4. **Business Logic:** Focus on WHY, not just WHAT the code does
5. **Modern Thinking:** Suggest improvements, not just replication
6. **Prioritization:** Help identify critical vs. nice-to-have features

## Resources

This skill includes bundled resources to support PRD generation:

### scripts/
- `analyze_form.py` - Analyzes VBA form files to extract structure, dependencies, and functionality
- `generate_flow.py` - Generates Mermaid navigation flow diagrams from analysis results

### references/
- `prd_templates.md` - Standard templates for master and module PRDs
- `migration_patterns.md` - Best practices for VBA-to-modern-web migration patterns

## Completion Checklist

Before finalizing PRDs, confirm:
- [ ] All forms have been reviewed
- [ ] Screenshots obtained for key forms
- [ ] Data model fully documented
- [ ] Navigation flows mapped
- [ ] User stories created for all major features
- [ ] API specifications defined
- [ ] Dead code identified
- [ ] Migration strategy outlined
- [ ] Tech stack recommended
- [ ] All PRD files generated

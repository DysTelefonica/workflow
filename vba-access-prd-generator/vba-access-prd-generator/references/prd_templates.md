# PRD Template Structure

This template defines the structure for Product Requirements Documents generated from VBA Access applications.

## Main PRD Structure (Master Document)

```markdown
# PRD: [Application Name]

## 1. Executive Summary
- Application purpose
- Target users
- Key business value
- Migration goals

## 2. Current System Overview
- Technology stack (VBA Access)
- Database architecture summary
- Key entities and relationships
- Integration points

## 3. Functional Areas
- List of major functional modules
- Links to detailed module PRDs

## 4. Data Model
- ERD diagram
- Table descriptions
- Key relationships
- Data volumes and growth

## 5. User Workflows
- Main user journeys
- Entry points
- Navigation flows

## 6. Non-Functional Requirements
- Performance requirements
- Security considerations
- Data migration needs
- Compliance requirements

## 7. Migration Strategy
- Phasing approach
- Risk assessment
- Dependencies
- Success criteria

## 8. Technical Debt & Code Quality
- Dead code identified
- Unused forms
- Refactoring opportunities
- Code complexity analysis

## 9. Appendices
- Links to detailed module PRDs
- Technical specifications
- API requirements
```

## Module/Form PRD Structure (Detailed Documents)

```markdown
# PRD Module: [Form/Module Name]

## 1. Overview
- Module purpose
- User role(s) that use it
- Business process supported

## 2. User Interface
### 2.1 Form Layout
- Screenshot/mockup
- Control inventory
- Layout description

### 2.2 User Interactions
- Available actions
- Validation rules
- Error handling

## 3. Functional Requirements
### 3.1 User Stories
- As a [user], I want to [action] so that [benefit]
- Acceptance criteria for each story

### 3.2 Business Rules
- Validation rules
- Calculation logic
- Workflow rules

## 4. Data Requirements
### 4.1 Data Sources
- Tables used
- Query patterns
- RecordSources

### 4.2 CRUD Operations
- Create operations
- Read operations
- Update operations
- Delete operations

## 5. Integration Points
### 5.1 Form Navigation
- Forms opened from this form
- Forms that open this form

### 5.2 Module Dependencies
- Shared modules used
- External APIs called

## 6. Technical Details
### 6.1 Code Structure
- Functions inventory
- Event handlers
- Global variables

### 6.2 SQL Queries
- Main queries used
- Performance considerations

## 7. Migration Considerations
### 7.1 REST API Endpoints Needed
- Endpoint specifications
- Request/response formats

### 7.2 Modern UI Equivalent
- Recommended web components
- UX improvements

### 7.3 Code Refactoring
- Code to eliminate
- Simplification opportunities
- Security improvements

## 8. Testing Considerations
- Key test scenarios
- Edge cases
- Data validation tests
```

## User Story Template

```markdown
### User Story: [Story Title]

**As a** [user role]
**I want to** [action/goal]
**So that** [business benefit]

**Acceptance Criteria:**
- [ ] Criterion 1
- [ ] Criterion 2
- [ ] Criterion 3

**Technical Notes:**
- VBA implementation: [description]
- Modern equivalent: [recommendation]

**Priority:** High/Medium/Low
```

## API Endpoint Specification Template

```markdown
### Endpoint: [Endpoint Name]

**Method:** GET/POST/PUT/DELETE
**Path:** `/api/v1/[resource]`

**Purpose:** [What this endpoint does]

**Request Parameters:**
```json
{
  "param1": "type",
  "param2": "type"
}
```

**Response:**
```json
{
  "data": {},
  "status": "success"
}
```

**Business Logic:**
- [Rule 1]
- [Rule 2]

**VBA Origin:** [Original form/function]
```

# Web Development Skill Matrix & Orchestration Protocol

## 1. Essential Skill Matrix
To ensure GreenPowerLink meets professional engineering standards, our development lifecycle requires these core competencies:

*   **Frontend (Delegated to @frontend-expert):**
    *   Responsive Design (Tailwind CSS, mobile-first).
    *   Component Architecture (Astro/React components).
    *   Performance Optimization (Core Web Vitals).
*   **Backend & Data (Delegated to @backend-expert):**
    *   API Design (RESTful, Type-safe endpoints).
    *   Database Management (Data modeling, integrity).
    *   Content Management (Markdown-driven content).
*   **Security (Delegated to @security-expert):**
    *   Vulnerability Assessment (XSS, CSRF, Injection protection).
    *   Secure Auth Flows (JWT, OAuth).
    *   Infrastructure Hardening.
*   **DevOps (Delegated to @devops-expert):**
    *   Infrastructure as Code (Proxmox, Nginx config).
    *   CI/CD Pipelines (Git-based automation).
    *   Static Hosting Optimization.

## 2. Manager Orchestration Protocol
The Manager (Primary Agent) MUST follow this protocol for every new task:

1.  **Decomposition**: Break the user's request into actionable items.
2.  **Delegation Check**:
    *   Does this involve UI or layout? -> Call **@frontend-expert**.
    *   Does this involve data or server logic? -> Call **@backend-expert**.
    *   Does this touch sensitive info or risk exposure? -> Call **@security-expert**.
    *   Does this impact hosting, Git, or server connectivity? -> Call **@devops-expert**.
3.  **Synthesis**: The Manager reviews the specialists' outputs, verifies compliance with the `100% Free Architecture` mandate, and merges the final solution.
4.  **Verification**: Before finalizing, the Manager ensures a commit and push to Git occurs.

---
*Failure to follow this protocol will result in a re-evaluation of the architectural path.*

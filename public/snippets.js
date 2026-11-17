/* SlideTeX — palette and examples. "¦" marks where the caret goes after insertion. */
window.SLIDETEX_SNIPPETS = [
  {
    tab: "Grec",
    items: [
      ["α", "\\alpha"], ["β", "\\beta"], ["γ", "\\gamma"], ["δ", "\\delta"], ["ε", "\\varepsilon"],
      ["ζ", "\\zeta"], ["η", "\\eta"], ["θ", "\\theta"], ["κ", "\\kappa"], ["λ", "\\lambda"],
      ["μ", "\\mu"], ["ν", "\\nu"], ["ξ", "\\xi"], ["π", "\\pi"], ["ρ", "\\rho"], ["σ", "\\sigma"],
      ["τ", "\\tau"], ["φ", "\\varphi"], ["χ", "\\chi"], ["ψ", "\\psi"], ["ω", "\\omega"],
      ["Γ", "\\Gamma"], ["Δ", "\\Delta"], ["Θ", "\\Theta"], ["Λ", "\\Lambda"], ["Π", "\\Pi"],
      ["Σ", "\\Sigma"], ["Φ", "\\Phi"], ["Ψ", "\\Psi"], ["Ω", "\\Omega"],
    ],
  },
  {
    tab: "Opérateurs",
    items: [
      ["a⁄b", "\\frac{¦}{}"], ["xⁿ", "^{¦}"], ["xₙ", "_{¦}"], ["√", "\\sqrt{¦}"], ["ⁿ√", "\\sqrt[¦]{}"],
      ["∑", "\\sum_{¦}^{}"], ["∏", "\\prod_{¦}^{}"], ["∫", "\\int_{¦}^{}"], ["∬", "\\iint"], ["∮", "\\oint"],
      ["lim", "\\lim_{¦ \\to }"], ["∂", "\\partial"], ["∇", "\\nabla"], ["∞", "\\infty"],
      ["±", "\\pm"], ["×", "\\times"], ["·", "\\cdot"], ["÷", "\\div"], ["∘", "\\circ"],
      ["d/dx", "\\frac{d¦}{dx}"], ["∂/∂x", "\\frac{\\partial ¦}{\\partial x}"], ["(n k)", "\\binom{¦}{}"],
    ],
  },
  {
    tab: "Relations",
    items: [
      ["≤", "\\leq"], ["≥", "\\geq"], ["≠", "\\neq"], ["≈", "\\approx"], ["≡", "\\equiv"], ["∝", "\\propto"],
      ["∼", "\\sim"], ["≪", "\\ll"], ["≫", "\\gg"], ["∈", "\\in"], ["∉", "\\notin"], ["⊂", "\\subset"],
      ["⊆", "\\subseteq"], ["∪", "\\cup"], ["∩", "\\cap"], ["∅", "\\emptyset"], ["∀", "\\forall"],
      ["∃", "\\exists"], ["¬", "\\neg"], ["∧", "\\land"], ["∨", "\\lor"],
      ["→", "\\to"], ["⇒", "\\Rightarrow"], ["⇔", "\\Leftrightarrow"], ["↦", "\\mapsto"],
    ],
  },
  {
    tab: "Accents",
    items: [
      ["x⃗", "\\vec{¦}"], ["x̂", "\\hat{¦}"], ["x̄", "\\bar{¦}"], ["x̃", "\\tilde{¦}"], ["ẋ", "\\dot{¦}"],
      ["ẍ", "\\ddot{¦}"], ["‾", "\\overline{¦}"], ["_", "\\underline{¦}"], ["⏞", "\\overbrace{¦}^{}"],
      ["⏟", "\\underbrace{¦}_{}"], ["ℝ", "\\mathbb{R}"], ["ℕ", "\\mathbb{N}"], ["ℤ", "\\mathbb{Z}"],
      ["ℂ", "\\mathbb{C}"], ["𝒜", "\\mathcal{¦}"], ["𝐀", "\\mathbf{¦}"], ["Aa", "\\text{¦}"],
      ["(…)", "\\left( ¦ \\right)"], ["[…]", "\\left[ ¦ \\right]"], ["{…}", "\\left\\{ ¦ \\right\\}"], ["|…|", "\\left| ¦ \\right|"],
    ],
  },
  {
    tab: "Structures",
    items: [
      ["( )₂ₓ₂", "\\begin{pmatrix} ¦ & \\\\  & \\end{pmatrix}"],
      ["[ ]₂ₓ₂", "\\begin{bmatrix} ¦ & \\\\  & \\end{bmatrix}"],
      ["|  |", "\\begin{vmatrix} ¦ & \\\\  & \\end{vmatrix}"],
      ["{ cas", "\\begin{cases} ¦ & \\text{si } \\\\  & \\text{sinon} \\end{cases}"],
      ["align", "\\begin{aligned} ¦ &= \\\\ &= \\end{aligned}"],
      ["⋯", "\\cdots"], ["⋮", "\\vdots"], ["⋱", "\\ddots"],
      ["⟨ ⟩", "\\langle ¦ \\rangle"], ["‖ ‖", "\\lVert ¦ \\rVert"], ["⌊ ⌋", "\\lfloor ¦ \\rfloor"],
      ["☐", "\\boxed{¦}"], ["⟶ᶠ", "\\xrightarrow{¦}"],
    ],
  },
];

window.SLIDETEX_EXAMPLES = [
  { name: "Gauss", mode: "display", src: "\\int_{-\\infty}^{+\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}" },
  { name: "Second degré", mode: "display", src: "x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}" },
  { name: "Euler", mode: "display", src: "e^{i\\pi} + 1 = 0" },
  { name: "Série", mode: "display", src: "\\sum_{n=1}^{\\infty} \\frac{1}{n^2} = \\frac{\\pi^2}{6}" },
  { name: "Matrice", mode: "display", src: "A = \\begin{pmatrix} a_{11} & \\cdots & a_{1n} \\\\ \\vdots & \\ddots & \\vdots \\\\ a_{m1} & \\cdots & a_{mn} \\end{pmatrix}" },
  { name: "Cas", mode: "display", src: "|x| = \\begin{cases} x & \\text{si } x \\geq 0 \\\\ -x & \\text{sinon} \\end{cases}" },
  { name: "Maxwell", mode: "display", src: "\\begin{aligned}\n\\nabla \\cdot \\mathbf{E} &= \\frac{\\rho}{\\varepsilon_0} \\\\\n\\nabla \\cdot \\mathbf{B} &= 0 \\\\\n\\nabla \\times \\mathbf{E} &= -\\frac{\\partial \\mathbf{B}}{\\partial t} \\\\\n\\nabla \\times \\mathbf{B} &= \\mu_0 \\mathbf{J} + \\mu_0 \\varepsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t}\n\\end{aligned}" },
  { name: "Schrödinger", mode: "display", src: "i\\hbar \\frac{\\partial}{\\partial t} \\Psi(\\mathbf{r},t) = \\left[ -\\frac{\\hbar^2}{2m}\\nabla^2 + V(\\mathbf{r},t) \\right] \\Psi(\\mathbf{r},t)" },
  { name: "Chimie (mhchem)", mode: "display", src: "\\ce{CO2 + C -> 2 CO}", preambleHint: "\\usepackage[version=4]{mhchem}" },
  { name: "Texte + maths", mode: "text", src: "Soit $f$ continue sur $[a,b]$. Alors $\\int_a^b f(x)\\,dx$ existe." },
  { name: "TikZ (LaTeX local)", mode: "text", engine: "latex", preambleHint: "\\usepackage{tikz}",
    src: "\\begin{tikzpicture}[scale=1.2]\n  \\draw[->] (-0.2,0) -- (3.2,0) node[right] {$x$};\n  \\draw[->] (0,-0.2) -- (0,2.2) node[above] {$y$};\n  \\draw[thick, domain=0:3, smooth, variable=\\x, red] plot ({\\x}, {0.2*\\x*\\x});\n  \\node at (2.2,1.6) {$y = \\frac{x^2}{5}$};\n\\end{tikzpicture}" },
];

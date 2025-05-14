import re
import logging

def check_latex():
    """Check if a LaTeX distribution is installed."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        from matplotlib import pyplot as plt
        plt.figure()
        plt.text(0.5, 0.5, r'$\alpha$', fontsize=12, ha='center', va='center')
        plt.axis('off')
        plt.savefig('test_latex.png', format='png', dpi=100, bbox_inches='tight')
        plt.close()
        import os
        if os.path.exists('test_latex.png'):
            os.remove('test_latex.png')
            logging.info("LaTeX distribution found and functional")
            return True
        logging.error("LaTeX rendering failed: no output file generated")
        return False
    except Exception as e:
        logging.error(f"LaTeX check failed: {e}")
        return False

def find_latex_equations(text):
    """Find LaTeX equations in text and return matches and cleaned equations."""
    if not text:
        logging.info("No text provided for LaTeX equation detection")
        return {'equations': [], 'matches': []}
    
    # Regex patterns for LaTeX equations
    patterns = [
        (r'\\\[(.*?)\\\]', True),           # Display math \[...\]
        (r'\\\((.*?)\\\)', False),         # Inline math \(...\)
        (r'\$\$(.*?)\$\$', True),          # Display math $$...$$
        (r'\$(.*?)\$', False),             # Inline math $...$
        (r'\\begin\{equation\}(.*?)\\end\{equation\}', True)  # Equation environment
    ]
    
    equations = []
    matches = []
    
    for pattern, is_display in patterns:
        for match in re.finditer(pattern, text, re.DOTALL):
            # Extract equation without delimiters
            equation = match.group(1).strip() if pattern not in [r'\\\((.*?)\\\)', r'\$(.*?)\$'] else match.group(1).strip()
            if equation:
                # Clean equation for rendering
                cleaned_equation = equation.replace('\n', ' ').strip()
                equations.append(cleaned_equation)
                matches.append({
                    'start': match.start(),
                    'end': match.end(),
                    'equation': cleaned_equation,
                    'is_display': is_display,
                    'raw_match': match.group(0)
                })
                logging.info(f"Detected equation: {cleaned_equation[:50]}... (start={match.start()}, end={match.end()})")
    
    # Sort matches by start position to maintain text order
    matches.sort(key=lambda x: x['start'])
    sorted_equations = [m['equation'] for m in matches]
    
    logging.info(f"Found {len(sorted_equations)} LaTeX equations")
    for i, eq in enumerate(sorted_equations, 1):
        logging.info(f"Equation {i}: {eq[:50]}...")
    
    return {
        'equations': sorted_equations,
        'matches': matches
    }
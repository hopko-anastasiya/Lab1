grammar Formula;

formula: expr EOF;

expr
    : expr '^' expr                  # PowerExpr
    | '-' expr                       # UnaryMinusExpr
    | '+' expr                       # UnaryPlusExpr
    | expr '*' expr                  # MulExpr
    | expr '/' expr                  # DivExpr
    | expr '+' expr                  # AddExpr
    | expr '-' expr                  # SubExpr
    | expr '=' expr                  # EqualExpr
    | expr '<>' expr                 # NotEqualExpr
    | expr '<' expr                  # LessExpr
    | expr '>' expr                  # GreaterExpr
    | expr '<=' expr                 # LessEqualExpr
    | expr '>=' expr                 # GreaterEqualExpr
    | '(' expr ')'                   # ParensExpr
    | CELL                           # CellExpr
    | NUMBER                         # NumberExpr
    ;

NUMBER      : DIGIT+(('.'|',')DIGIT+)?;
CELL        : LETTER+DIGIT+;
 
DIGIT       : [0-9];
LETTER      : [A-Z];

WS          : [ \t\r\n]+ -> skip ; 
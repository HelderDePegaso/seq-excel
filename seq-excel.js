class SeqExcel {
    static __literalLimit = "XFD"
    static __numberLimit = 1048576

    static nextValue(comb) {
        if(!comb) return "A1";

        let splitedComb = splitComb(comb)

        if(splitedComb == undefined) return undefined // Se teve um Erro na separção

        if(splitedComb.literals == SeqExcel.__literalLimit  && splitedComb.numbers == SeqExcel.__numberLimit) return undefined // Se o limite foi atingido

        let nextLiteral = undefined
        let nextNumber = undefined

        if(splitedComb.literals == SeqExcel.__literalLimit  && splitedComb.numbers < SeqExcel.__numberLimit) {
            nextNumber = Number(splitedComb.numbers) + 1
            nextLiteral = "A";
        } else if(splitedComb.literals != SeqExcel.__literalLimit) {
            nextNumber = splitedComb.numbers ;
            nextLiteral = nextCollum(splitedComb.literals);
        }
        
        let newComb =  nextLiteral + nextNumber

        return newComb;
                            

    }
}


function splitComb(comb) {
    // Se o tamanho da string for inválido
    if(!(comb.length >= 2) && 
       !(comb.length <= String(SeqExcel.__numberLimit).length + SeqExcel.__literalLimit.length)) // Se o tamanho da string não for >= 2 e <= 10
            return undefined

    if(!isInSet(comb)) return undefined
    
    let obj = {literals : "", numbers : ""}
    for (let i = 0; i < comb.length; i++) {
        const caractere = comb[i];
        if(typeChar(caractere) == 0) obj.literals += caractere
        if(typeChar(caractere) == 1) obj.numbers += caractere
    }

    if(obj.literals != "" && obj.numbers != "") return obj
    else return undefined
}

function isInSet(params) {
    if(/^[A-Z]+[0-9]+$/.test(params)) return true
    else return false
}


function typeChar(caractere) {
    // Utiliza uma expressão regular para verificar se o caractere é uma letra
    if( /[a-zA-Z]/.test(caractere)) return 0;
    // Utiliza uma expressão regular para verificar se o caractere é um número (d/)
    else if(/^\d+$/.test(caractere)) return 1; 
}

function positionBefore(char1, char2) {
    let alfabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    let posChar1 = alfabet.indexOf(char1)
    let posChar2 = alfabet.indexOf(char2)

    return (posChar1 < posChar2) ? -1 : (posChar1 == posChar2) ? 0 : 1
    
}

function nextLetter(char) {
    let alfabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    let posChar = alfabet.indexOf(char) + 1
    if(posChar == alfabet.length) return "A"
    else return alfabet[posChar]
}

function nextCollum(coluna) {
    if (coluna === "XFD") {
        return "XFD";  // Alcançou o limite máximo
    }

    if(coluna === "Z") {
        return "AA";  
    }

    const proximoCodigo = (codigo) => (codigo === 90) ? 65 : codigo + 1;  // 90 é o código ASCII para 'Z'

    const codigoAtual = coluna.charCodeAt(coluna.length - 1);
    const proximoCodigoFinal = proximoCodigo(codigoAtual);

    if (proximoCodigoFinal === 65) {
        // Se o próximo código for 'A', incrementa a coluna anterior
        const colunaAnterior = coluna.slice(0, -1);
        return nextCollum(colunaAnterior) + "A";
    } else {
        // Caso contrário, apenas incrementa a última letra
        let sai = coluna.slice(0, -1) + String.fromCharCode(proximoCodigoFinal);
        return (sai.startsWith("\x00")) ? sai.slice(1) : sai
         
    }
}
function nextLiteral3(currentLiteral) {
    const nextLetter = (letter) => letter === "Z" ? "A" : String.fromCharCode(letter.charCodeAt(0) + 1);

    if (currentLiteral === "XFD") {
        return "XFD";
    }

    const [first, second, third] = currentLiteral;

    if (first === "Z" && second === "Z") {
        return "AAA";
    }

    if (third !== "D") {
        return first + second + nextLetter(third);
    }

    if (second !== "F") {
        return first + second + nextLetter(third);
    }

    if (first !== "X") {
        return first + second + nextLetter(third);
    }

    return "XFD";
}


module.exports = SeqExcel

//
/*
(splitedComb.literals.length == 1 && splitedComb.literals.endsWith("Z")) ? "AA" : 
                          (splitedComb.literals.length == 2 && splitedComb.literals.endsWith("Z") && splitedComb.literals[0] != "Z") ? nextLetter(splitedComb.literals[0]) +"A" :
                          (splitedComb.literals == "ZZ") ? "AAA" : 

                           (splitedComb.literals.length == 3 && splitedComb.literals != "XFD" ) 
                           ? 
                                (splitedComb.literals.length == 3 && positionBefore(splitedComb.literals[2], "D") == -1) 
                                ? splitedComb.literals[0] + splitedComb.literals[1]+ nextLetter(splitedComb.literals[2]) 
                                : (splitedComb.literals.length == 3 && positionBefore(splitedComb.literals[2], "D") == 0) 
                                    ? (splitedComb.literals.length == 3 && positionBefore(splitedComb.literals[1], "F") == -1) 
                                        ?  splitedComb.literals[0] + splitedComb.literals[1] + nextLetter(splitedComb.literals[2]) 
                                        :  (splitedComb.literals.length == 3 && positionBefore(splitedComb.literals[1], "F") ==  0) 
                                        ?  (splitedComb.literals.length == 3 && positionBefore(splitedComb.literals[0], "X") ==  -1) 
                                           ? splitedComb.literals[0] + splitedComb.literals[1] + nextLetter(splitedComb.literals[2]) 
                                           : splitedComb.literals[0] + splitedComb.literals[1] + splitedComb.literals[2]                                                                     
                                        : "XFD"
                                    : "XFD"
                            : "XFD";*/
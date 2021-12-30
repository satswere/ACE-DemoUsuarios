(error, messages: any, rawResponse?: any) => {
  
    console.log("necesitamos saber si aqui se puede ver la prueba");

    let Usuarios = messages.value ;
    let nuevoValor = []
    Usuarios.forEach(element => {
      const {displayName,jobTitle,mail} = element
      nuevoValor.push({ Name: displayName, job: jobTitle, mail: mail })
    });
    console.log(nuevoValor);
    

///     console.log(messages.value[2]);
   // console.log(messages.value[2].mail);

  }
using System.Collections.Generic;

namespace BlazorFiles.Api.TablasParametricasDto
{
    public class UpdateTablasParametricasResponseDto
    {
        public IEnumerable<OrganismosDto> ListOrgnismos { get; set; }
        public IEnumerable<MarcasDto> ListMarcas { get; set; }
        //Aqui se colocan todas las listas de las tablas 
        //paramétricas que se quieren retornar
    }
}

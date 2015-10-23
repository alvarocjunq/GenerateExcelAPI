package br.com.gexapi.data;

import br.com.gexapi.main.PositionExcel;

public class Cliente {
	private Integer id;
	private String nome;
	
	public Cliente(Integer id, String nome){
		this.id = id;
		this.nome = nome;
	}
	
	@PositionExcel(posicao={0})
	public Integer getId() { return id; }
	
	@PositionExcel(posicao={1})
	public String getNome() { return nome; }
}

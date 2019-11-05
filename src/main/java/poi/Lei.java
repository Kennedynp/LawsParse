package poi;

public class Lei {
	
	private String chave;
	private String chaveArtigo;
	private String descricao;
		
	public Lei(String chave, String chaveArtigo, String descricao) {
		super();
		this.chave = chave;
		this.chaveArtigo = chaveArtigo;
		this.descricao = descricao;
	}

	public String getChave() {
		return chave;
	}
	
	public String getChaveArtigo() {
		return chaveArtigo;
	}
	
	public String getDescricao() {
		return descricao;
	}	
		
}

package br.com.gexapi;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
/**
 * Annotation 
 * @author Alvaro Junqueira
 *
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface PositionExcel {

	/**
	 * @return Array de inteiros, iniciando do 0 (zero) para definir em qual coluna vai aparecer no Excel
	 */
	int[] posicao();
}
-- Central de Credenciamento FAF: leitura dos jogos do Urano.
-- Só LEITURA: as funções devolvem dados públicos (competições, jogos, clubes,
-- escudos, estádios) e não permitem criar, alterar nem apagar nada no Urano.

create or replace function public.credenciamento_competicoes()
returns table (id text, nome text, temporada integer, categoria text, logo text)
language sql stable security definer set search_path = public as $$
  select c.id, c.name, c.season, c.category, c.logo
  from competitions c
  where c.active and c.deleted_at is null
  order by c.season desc, c.name;
$$;

create or replace function public.credenciamento_jogos(p_competicao text)
returns table (
  id text, rodada text, fase text, data text, hora text,
  mandante_id text, mandante text, escudo_mandante text,
  visitante_id text, visitante text, escudo_visitante text, estadio text
)
language sql stable security definer set search_path = public as $$
  select m.id, m.round, m.phase, m.date, m.time,
         -- escudo só quando for link; os do Urano são arquivos assets/escudos/<id>.png
         m.home_club_id, coalesce(h.short_name, m.home_club_id), case when h.shield like 'http%' then h.shield end,
         m.away_club_id, coalesce(a.short_name, m.away_club_id), case when a.shield like 'http%' then a.shield end,
         s.name
  from matches m
  join competitions c on c.id = m.competition_id and c.active and c.deleted_at is null
  left join clubs h on h.id = m.home_club_id
  left join clubs a on a.id = m.away_club_id
  left join stadiums s on s.id = m.stadium_id
  where m.competition_id = p_competicao;
$$;

-- Ninguém além destes dois papéis pode chamar; e só chamar (nada de escrita).
revoke all on function public.credenciamento_competicoes() from public;
revoke all on function public.credenciamento_jogos(text) from public;
grant execute on function public.credenciamento_competicoes() to anon, authenticated;
grant execute on function public.credenciamento_jogos(text) to anon, authenticated;

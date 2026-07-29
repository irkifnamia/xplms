-- XPLMS migration 005: classify study materials by chapter and resource type

begin;

alter table public.materials
  add column if not exists chapter integer,
  add column if not exists material_type text;

alter table public.materials
  drop constraint if exists materials_chapter_check,
  add constraint materials_chapter_check
    check (chapter in (1, 2, 5, 8, 9, 10)),
  drop constraint if exists materials_material_type_check,
  add constraint materials_material_type_check
    check (
      material_type in (
        'Infographic',
        'Notes',
        'Exercise',
        'Extra',
        'Reference',
        'Other'
      )
    );

create index if not exists materials_chapter_type_idx
  on public.materials (chapter, material_type, uploaded_at desc);

commit;
